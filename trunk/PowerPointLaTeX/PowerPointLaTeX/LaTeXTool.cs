#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;

namespace PowerPointLaTeX
{
    public enum EquationType
    {
        None,
        // has inline LaTeX codes (but not compiled)
        HasInlines,
        // has compiled LaTeX codes
        HasCompiledInlines,
        // a compiled LaTeX code element (picture)
        Inline,
        // an equation (picture)
        Equation,
    }

    static class EquationTypeShapeExtension
    {
        public static bool IsEquation(this Shape shape)
        {
            return shape.LaTeXTags().Type == EquationType.Equation;
        }
    }

    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool : PowerPointLaTeX.ILaTeXTool
    {
        private const char NoneBreakingSpace = (char) 8201;
        private const int InitialEquationFontSize = 44;

        private Microsoft.Office.Interop.PowerPoint.Application Application
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        internal Presentation ActivePresentation
        {
            get { return Application.ActivePresentation; }
        }

        internal Slide ActiveSlide
        {
            get { return Application.ActiveWindow.View.Slide as Slide; }
        }

        /// <summary>
        /// returns whether the addin is enabled in the current context (ie presentation)
        /// but is also affected by the global addin settings, of course.
        /// </summary>
        internal bool AddInEnabled
        {
            get
            {
                return !ActivePresentation.Final && Properties.Settings.Default.EnableAddIn && Compatibility.IsSupportedPresentation( ActivePresentation );
            }
        }

        /// <returns>A valid picture shape from the data or null if creation failed</returns>
        private Shape CreatePictureShapeFromImage(Slide slide, Image image)
        {           
            IDataObject oldClipboardContent = null;
            try {
                oldClipboardContent = Clipboard.GetDataObject();
            }
            catch { Debug.Assert( false, "Retrieving the current clipboard contents failed!"); }

            Clipboard.SetImage(image);

            ShapeRange pictureRange = slide.Shapes.Paste();
            if( oldClipboardContent != null )
                Clipboard.SetDataObject(oldClipboardContent);

            if( pictureRange == null ) {
                return null;
            }

            Trace.Assert(pictureRange.Count == 1);
            return pictureRange[1];
        }

        private Shape GetPictureShapeFromLaTeXCode( Slide currentSlide, string latexCode, float fontSize )
        {
            // check the cache first
            int baselineOffset;
            float wantedPixelsPerEmHeight = GetPixelsPerEmHeight( fontSize, WindowsDPISetting );
            float actualPixelsPerEmHeight = GetPixelsPerEmHeight( fontSize, RenderDPISetting );
            Image image = GetImageForLaTeXCode( latexCode, ref actualPixelsPerEmHeight, out baselineOffset );
            if (image == null) {
                return null;
            }

            Shape picture = CreatePictureShapeFromImage(currentSlide, image);
            if( picture == null ) {
                return null;
            }

            // make white the transparent color
            picture.PictureFormat.TransparencyColor = ~0;
            picture.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoCTrue;

            picture.AlternativeText = latexCode;

            // prescale the image to the wanted em height here
            float scaleRatio = wantedPixelsPerEmHeight / actualPixelsPerEmHeight;

            picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            picture.Height *= scaleRatio;
            picture.Width *= scaleRatio;

            // store the baseline offset as percentage value instead of pixels to support rescaling the image
            picture.LaTeXTags().BaseLineOffset.value = (float) baselineOffset / image.Height;
            //picture.LaTeXTags().PixelsPerEmHeight = actualPixelsPerEmHeight;
            string shortenedName;
            if (latexCode.Length > 32) {
                shortenedName = latexCode.Substring(0, 32) + "..";
            }
            else {
                shortenedName = latexCode;
            }
            picture.Name = "LaTeX: " + shortenedName;

            return picture;
        }

        // TODO: move GetPixelsPerEmHeight and WindowsDPISetting into a helper class? [5/23/2010 Andreas]
        public const int WindowsDPISetting = 96;
        // renders all formulas at a higher resolution than necessary to allow for zooming
        public const int RenderDPISetting = 300;
        public const int PrintPtsPerInch = 72;

        public static float GetPixelsPerEmHeight( float fontSizeInPoints, int targetPixelsPerInch ) {
            return fontSizeInPoints / PrintPtsPerInch * targetPixelsPerInch;
        }

        public static float PixelHeightToFontSize(float pixelHeight) {
            return pixelHeight / WindowsDPISetting * PrintPtsPerInch;
        }

        public static float FontSizeToPixelHeight(float fontSize) {
            return fontSize / PrintPtsPerInch * WindowsDPISetting;
        }

        /// <summary>
        /// Compile latexCode into an inline shape
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="textShape"></param>
        /// <param name="latexCode"></param>
        /// <param name="codeRange"></param>
        /// <returns></returns>
        private Shape CompileInlineLaTeXCode(Slide slide, Shape textShape, string latexCode, TextRange codeRange)
        {
            Shape picture = GetPictureShapeFromLaTeXCode( slide, latexCode, codeRange.Font.Size );
            if (picture == null)
            {
                return null;
            }

            // add tags to the picture
            picture.LaTeXTags().Code.value = latexCode;
            picture.LaTeXTags().Type.value = EquationType.Inline;
            picture.LaTeXTags().LinkID.value = textShape.Id;
            
            float baselineOffset = picture.LaTeXTags().BaseLineOffset;
            float heightInPts = PixelHeightToFontSize( picture.Height );

            FontFamily fontFamily;
            try
            {
                fontFamily = new FontFamily(codeRange.Font.Name);
            }
            catch( Exception exception ) {
                // TODO: add message box and inform the user about it [9/20/2010 Andreas]
                MessageBox.Show("Failed to load font information (using Times New Roman as substitute). Error: " + exception, "PowerPoint LaTeX");
                fontFamily = new FontFamily("Times New Roman");
            }

            // from top to baseline
            float ascentHeight = (float) (codeRange.Font.Size * ((float) fontFamily.GetCellAscent( FontStyle.Regular ) / fontFamily.GetEmHeight( FontStyle.Regular )));
            float descentRatio = (float) fontFamily.GetCellDescent( FontStyle.Regular ) / fontFamily.GetEmHeight( FontStyle.Regular );
            float descentHeight = (float) (codeRange.Font.Size * descentRatio);

            float ascentSize = Math.Max( 0, 1 - baselineOffset ) * heightInPts;
            float descentSize = Math.Max( 0, baselineOffset ) * heightInPts;
            
            float factor = 1.0f;
            if( ascentSize > 0 ) {
                factor = Math.Max( factor, ascentSize / ascentHeight );
            }
            if( descentSize > 0 ) {
                factor = Math.Max( factor, descentSize / descentHeight );
            }

            // dont let it get too big (starts to look ridiculous otherwise)
            if( factor <= 1.5f ) {
                codeRange.Font.Size *= factor;
            }
            else {
                // keep linespacing intact (assuming that the line spacing scales with the font height)
                float lineSpacing = (float) (codeRange.Font.Size * ((float) fontFamily.GetLineSpacing( FontStyle.Regular ) / fontFamily.GetEmHeight( FontStyle.Regular )));
                // additional line spacing
                codeRange.Font.Size *= (lineSpacing + 5.0f - codeRange.Font.Size + heightInPts) / lineSpacing;
                // just ignore the baseline offset
                picture.LaTeXTags().BaseLineOffset.value = descentRatio;
            }


            /*
            if( Math.Abs(baselineOffset) > 1 ) {
                            if( baselineOffset < -1 ) {
                                codeRange.Font.Size = heightInPts * (1 - baselineOffset); // ie 1 + abs( baselineOffset)
                                codeRange.Font.BaselineOffset = 1;
                            }
                            else / * baselineOffset > 1 * / {
                                codeRange.Font.Size = heightInPts * baselineOffset;
                                codeRange.Font.BaselineOffset = -1;
                           }
                        }
                        else {
                            // change the font size to keep the formula from overlapping with regular text (nifty :))
                            codeRange.Font.Size = heightInPts;
            
                            // BaseLineOffset > for subscript but PPT uses negative values for this
                            codeRange.Font.BaselineOffset = -baselineOffset;
                        }*/
            

            // disable word wrap
            codeRange.ParagraphFormat.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            // fill the text up with none breaking space to make it "wrap around" the formula
            FillTextRange(codeRange, NoneBreakingSpace, picture.Width);

            CopyInlineEffects( slide, textShape, codeRange, picture );

            return picture;
        }

        private int GetSafeEffectParagraph( Effect effect ) {
            try {
                return effect.Paragraph;
            }
            catch {
                return 1;
            }
        }

        private void CopyInlineEffects( Slide slide, Shape textShape, TextRange codeRange, Shape picture ) {
            try {
                // copy animations from the parent textShape
                Sequence sequence = slide.TimeLine.MainSequence;
                var effects =
                    from Effect effect in sequence
                    where effect.Shape.SafeThis() != null && effect.Shape == textShape &&
                    ((effect.EffectInformation.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph &&
                        Helpers.ParagraphContainsRange( textShape, GetSafeEffectParagraph( effect ), codeRange ))
                        || effect.EffectInformation.BuildByLevelEffect == MsoAnimateByLevel.msoAnimateLevelNone)
                    select effect;

                CopyEffectsTo( picture, true, sequence, effects );
            }
            catch {
                Debug.Fail( "CopyInlineEffects failed!" );
            }
        }

        private static void CopyEffectsTo(Shape target, bool setToWithPrevious, Sequence sequence, IEnumerable<Effect> effects)
        {
            foreach (Effect effect in effects)
            {
                int index = effect.Index + 1;
                Effect formulaEffect = sequence.Clone(effect, index);
                try
                {
                    formulaEffect = sequence.ConvertToBuildLevel(formulaEffect, MsoAnimateByLevel.msoAnimateLevelNone);
                }
                catch { }
                //formulaEffect = sequence.ConvertToTextUnitEffect(formulaEffect, MsoAnimTextUnitEffect.msoAnimTextUnitEffectMixed);
                if (setToWithPrevious)
                    formulaEffect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                try
                {
                    formulaEffect.Paragraph = 0;
                }
                catch { }
                formulaEffect.Shape = target;
                // Effect formulaEffect = sequence.AddEffect(picture, effect.EffectType, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, index);
            }
        }

        private static void FillTextRange(TextRange range, char character, float minWidth)
        {
            range.Text = character.ToString() + ( (char) 160 ).ToString(); // ;

            // line-breaks are futile, so break if one happens 
            float oldHeight = range.BoundHeight;
            while (range.BoundWidth < minWidth && oldHeight == range.BoundHeight)
            {
                range.Text += character.ToString() + ((char) 160).ToString();
            }
            if( oldHeight != range.BoundHeight ) {
                range.Text = range.Text.Remove( range.Text.Length - 2, 2 );
                range.Text += ((char)8232).ToString(); // new line that doesnt begin new paragraph
            }
        }

        private static void AlignFormulaWithText(TextRange codeRange, Shape picture)
        {
            // interesting fact: text filled with (at most one line of none-breaking) spaces -> BoundHeight == EmSize
            //codeRange.Text = " ";
            float fontHeight = codeRange.BoundHeight;
            FontFamily fontFamily = new FontFamily(codeRange.Font.Name);
            // from top to baseline
            float baselineHeight = (float) (fontHeight * ((float) fontFamily.GetCellAscent(FontStyle.Regular) / fontFamily.GetLineSpacing(FontStyle.Regular)));

            picture.Left = codeRange.BoundLeft;

            // DISABLED to try baseline feature from the miktex service [5/21/2010 Andreas]
 /*
            if (baselineHeight >= picture.Height)
             {
                 // baseline: center (assume that its a one-line codeRange)
                 picture.Top = codeRange.BoundTop + (baselineHeight - picture.Height) * 0.5f;
             }
             else
             {
                 // center the picture directly
                 picture.Top = codeRange.BoundTop + (fontHeight - picture.Height) * 0.5f;
             }*/
            picture.Top = codeRange.BoundTop + baselineHeight - (1.0f - picture.LaTeXTags().BaseLineOffset) * picture.Height;
 
        }

        /// <summary>
        /// Get the raw image data for some latex code or null if the compilation failed.
        /// It seems to be possible to receive a byte[0] array for some reason.
        /// </summary>
        /// <param name="latexCode"></param>
        /// <returns></returns>
        private byte[] GetImageDataForLaTeXCode(string latexCode, ref float pixelsPerEmHeight, out int baselineOffset)
        {
            byte[] imageData;

            byte[] cachedImageData = null;
            float cachedPixelsPerEmHeight = 0;
            int cachedBaselineOffset = 0;
            // TODO: rewrite the cache system to work even if the main thread is blocked [8/4/2009 Andreas]
            if (ActivePresentation.CacheTags()[latexCode].IsCached())
            {
                ActivePresentation.CacheTags()[ latexCode ].Query( out cachedImageData, out cachedPixelsPerEmHeight, out cachedBaselineOffset );

                // make sure we return a some-what meaningful array
                if( cachedImageData != null && cachedImageData.Length == 0 ) {
                    cachedImageData = null;
                }
            }

            // convert to int to avoid floating point issues [5/24/2010 Andreas]
            if( cachedImageData != null && (int) cachedPixelsPerEmHeight >= (int) pixelsPerEmHeight ) {
                // we can use the cached formula
                imageData = cachedImageData;
                pixelsPerEmHeight = cachedPixelsPerEmHeight;
                baselineOffset = cachedBaselineOffset;

                ActivePresentation.CacheTags()[ latexCode ].Use();
            }
            else {
                // try to render the formula using our LaTeX service
                Globals.ThisAddIn.LaTeXServices.Service.RenderLaTeXCode(latexCode, out imageData, ref pixelsPerEmHeight, out baselineOffset);

                if( imageData != null && imageData.Length > 0 ) {
                    // looks good, so cache it
                    ActivePresentation.CacheTags()[ latexCode ].Store( imageData, pixelsPerEmHeight, baselineOffset );
                }
                else {
                    // if this failed, use the result from the cache, can't be off worse
                    imageData = cachedImageData;
                    pixelsPerEmHeight = cachedPixelsPerEmHeight;
                    baselineOffset = cachedBaselineOffset;

                    if( cachedImageData != null ) {
                        ActivePresentation.CacheTags()[ latexCode ].Use();
                    }
                }
            }

            return imageData;
        }


        private Image GetImageFromImageData(byte[] imageData) {
            if( imageData == null ) {
                return null;
            }

            MemoryStream stream = new MemoryStream(imageData);
            Image image;
            try {
                image = Image.FromStream(stream);
            }
            catch {
                return null;
            }
            return image;
        }

        public Image GetImageForLaTeXCode(string latexCode, ref float pixelsPerEmHeight, out int baselineOffset) {
            byte[] imageData = GetImageDataForLaTeXCode( latexCode, ref pixelsPerEmHeight, out baselineOffset );
            return GetImageFromImageData(imageData);
        }

        private static bool IsEscapeCode(string code)
        {
            return code == "!";
        }

        private void CompileInlineTextRange(Slide slide, Shape shape, TextRange range)
        {
            int startIndex = 0;

            int codeCount = 0;

            List<TextRange> pictureRanges = new List<TextRange>();
            List<Shape> pictures = new List<Shape>();

            while (true)
            {
                bool inlineMode;
                int latexCodeStartIndex;
                int latexCodeEndIndex;
                int endIndex;

                int inlineStartIndex, displaystyleStartIndex;
                inlineStartIndex = range.Text.IndexOf("$$", startIndex);
                displaystyleStartIndex = range.Text.IndexOf( "$$[", startIndex );
                if( displaystyleStartIndex != -1 && displaystyleStartIndex <= inlineStartIndex ) {
                    inlineMode = false;

                    startIndex = displaystyleStartIndex;
                    latexCodeStartIndex = startIndex + 3;
                    latexCodeEndIndex = range.Text.IndexOf( "]$$", latexCodeStartIndex );
                    endIndex = latexCodeEndIndex + 3;
                } else if( inlineStartIndex != -1 ) {
                    inlineMode = true;

                    startIndex = inlineStartIndex;
                    latexCodeStartIndex = startIndex + 2;
                    latexCodeEndIndex = range.Text.IndexOf( "$$", latexCodeStartIndex );
                    endIndex = latexCodeEndIndex + 2;
                }
                else {
                    break;
                }

                if( latexCodeEndIndex == -1 )
                {
                    break;
                }

                int length = endIndex - startIndex;

                int latexCodeLength = latexCodeEndIndex - latexCodeStartIndex;
                string latexCode = range.Text.Substring( latexCodeStartIndex, latexCodeLength );
                // TODO: move this into its own function [1/5/2009 Andreas]
                latexCode = latexCode.Replace((char) 8217, '\'');
                // replace weird unicode - (hypens) with minus
                latexCode = latexCode.Replace( (char) 8208, '-' );
                latexCode = latexCode.Replace( (char) 8211, '-' );
                latexCode = latexCode.Replace( (char) 8212, '-' );
                latexCode = latexCode.Replace( (char) 8722, '-' );
                latexCode = latexCode.Replace( (char) 8209, '-' );
                latexCode = latexCode.Replace( (char) 8259, '-' );
                // replace ellipses with ...
                latexCode = latexCode.Replace( ((char) 8230).ToString(), "..." );

                // must be [[ then
                if( !inlineMode ) {
                    latexCode = @"\displaystyle{" + latexCode + "}";
                }

                LaTeXEntry tagEntry = shape.LaTeXTags().Entries[codeCount];
                tagEntry.Code.value = latexCode;
                // TODO: cohesion? [5/2/2009 Andreas]
                // save the font size because it might be changed later
                // +1 because IndexOf is base 0, but Characters uses base 1
                tagEntry.FontSize.value = range.Characters( latexCodeStartIndex + 1, latexCodeLength ).Font.Size;

                // escape $$!$$
                // +1 because IndexOf is base 0, but Characters uses base 1
                TextRange codeRange = range.Characters(startIndex + 1, length);
                if (!IsEscapeCode(latexCode))
                {
                    Shape picture = CompileInlineLaTeXCode(slide, shape, latexCode, codeRange);
                    if (picture != null)
                    {
                        tagEntry.ShapeId.value = picture.Id;

                        pictures.Add(picture);
                        pictureRanges.Add(codeRange);
                    }
                    else
                    {
                        codeRange.Text = "$Formula Error$";
                    }
                }
                else
                {
                    codeRange.Text = "$$";
                }

                tagEntry.StartIndex.value = codeRange.Start;
                tagEntry.Length.value = codeRange.Length;

                // NOTE: endIndex isnt valid anymore since we've removed some text [5/24/2010 Andreas
                // IndexOf uses base0, codeRange base1 => -1
                startIndex = codeRange.Start + codeRange.Length - 1;
                codeCount++;
            }

            // TODO: this doesn't work - simply disable autofit instead.. [1/5/2009 Andreas]
            // TODO: can we automate this? [2/26/2009 Andreas]
            /*
                   (new Thread( delegate() {
                            Thread.Sleep(100);
                            for (int i = 0; i < pictures.Count; i++)
                            {
                                TextRange codeRange = pictureRanges[i];
                                AlignFormulaWithText(codeRange, pictures[i]);
                            }
                        } )).Start();*/

            // now that everything has been converted we can position the formulas (pictures) in the text area
            for (int i = 0; i < pictures.Count; i++)
            {
                TextRange codeRange = pictureRanges[i];
                AlignFormulaWithText(codeRange, pictures[i]);
            }

            // update the type, too
            shape.LaTeXTags().Type.value = codeCount > 0 ? EquationType.HasCompiledInlines : EquationType.None;
        }

        public void CompileShape(Slide slide, Shape shape)
        {
            // we don't need to compile already compiled shapes (its also sensible to avoid destroying escape sequences or overwrite entries, etc.)
            // don't try to compile equations (or their sources) either
            EquationType type = shape.LaTeXTags().Type;
            if (type == EquationType.HasCompiledInlines || type == EquationType.Equation)
            {
                return;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    CompileInlineTextRange(slide, shape, textFrame.TextRange);
                }
            }
        }

        private void DecompileTextRange(Slide slide, Shape shape, TextRange range)
        {
            // make sure this is always valid, otherwise the code will do stupid things
            Debug.Assert(shape.LaTeXTags().Type == EquationType.HasCompiledInlines);

            LaTeXEntries entries = shape.LaTeXTags().Entries;
            int length = entries.Length;
            for (int i = length - 1; i >= 0; i--)
            {
                LaTeXEntry entry = entries[i];
                string latexCode = entry.Code;

                if (!IsEscapeCode(latexCode))
                {
                    int shapeID = entry.ShapeId;
                    // find the shape
                    Shape picture = slide.Shapes.FindById(shapeID);

                    //Debug.Assert(picture != null);
                    // fail gracefully
                    if (picture != null)
                    {
                        Debug.Assert(picture.LaTeXTags().Type == EquationType.Inline);
                        picture.Delete();
                    }

                    // release the cache entry, too
                    ActivePresentation.CacheTags()[latexCode].Release();
                }

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                if( latexCode.StartsWith( @"\displaystyle{" ) && latexCode.EndsWith( "}" ) ) {
                    codeRange.Text = "$$[" + latexCode.Substring( @"\displayStyle{".Length, latexCode.Length - 1 - @"\displayStyle{".Length ) + "]$$";
                }
                else {
                    codeRange.Text = "$$" + latexCode + "$$";
                }
                if (entry.FontSize != 0) {
                    codeRange.Font.Size = entry.FontSize;
                }
                codeRange.Font.BaselineOffset = 0.0f;
            }

            entries.Clear();
            shape.LaTeXTags().Type.value = EquationType.HasInlines;
        }

        public void DecompileShape(Slide slide, Shape shape)
        {
            // we don't need to decompile already shapes that aren't compiled
            if (shape.LaTeXTags().Type != EquationType.HasCompiledInlines)
            {
                return;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    DecompileTextRange(slide, shape, textFrame.TextRange);
                }
            }
        }

        /*
          private bool presentationNeedsCompile;
                 // TODO: use an exception, etc. to early-out [1/2/2009 Andreas]
                 private void ShapeNeedsCompileWalker(Slide slide, Shape shape) {
                     EquationType type = shape.LaTeXTags().Type;
                     if( type == EquationType.HasInlines ) {
                         presentationNeedsCompile = true;
                     }
                     if( type == EquationType.None ) {
                         // check it and assign a type if necessary
                         if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                         {
                             TextFrame textFrame = shape.TextFrame;
                             if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                             {
                                 // search for a $$ - if there is one occurrence, it needs to be compiled
                                 if(textFrame.TextRange.Text.IndexOf("$$") != -1) {
                                     presentationNeedsCompile = true;
                                     // update the type, too
                                     shape.LaTeXTags().Type.value = EquationType.HasInlines;
                                 } else {
                                     shape.LaTeXTags().Type.value = EquationType.None;
                                 }
                             }
                         }
                     }
                 }
        
                 public bool NeedsCompile(Presentation presentation) {
                     presentationNeedsCompile = false;
                     WalkPresentation(presentation, ShapeNeedsCompileWalker);
                     return presentationNeedsCompile;
                 }*/


        private void FinalizeShape(Slide slide, Shape shape)
        {
            CompileShape(slide, shape);
            shape.LaTeXTags().Clear();
        }

        

        public void CompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, CompileShape);
        }

        public void DecompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, DecompileShape);
        }

        public void CompilePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, CompileShape);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        public void FinalizePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, FinalizeShape);
            // purge the cache, too
            presentation.CacheTags().PurgeAll();
            presentation.SettingsTags().Clear();
        }

        public Shape CreateEmptyEquation()
        {
            const float width = 100, height = 60;

            Shape shape = ActiveSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 100, 100, width, height);
            shape.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorBackground1;
            shape.Fill.Solid();

            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            LaTeXTags tags = shape.LaTeXTags();
            tags.Type.value = EquationType.Equation;
            tags.OriginalWidth.value = width;
            tags.OriginalHeight.value = height;
            tags.FontSize.value = InitialEquationFontSize;

            return shape;
        }

        public Shape EditEquation( Shape equation, out bool cancelled ) {
            EquationEditor editor = new EquationEditor( equation.LaTeXTags().Code, equation.LaTeXTags().FontSize );
            DialogResult result = editor.ShowDialog();
            if( result == DialogResult.Cancel ) {
                cancelled = true;
                // don't change anything
                return equation;
            }
            else {
                cancelled = false;
            }

            // recompile the code
            //equation.TextFrame.TextRange.Text = equationSource.TextFrame.TextRange.Text;
            string latexCode = editor.LaTeXCode;

            Slide slide = equation.GetSlide();
            if (slide == null) {
                // TODO: what do we do in this case? [3/3/2009 Andreas]
                return equation;
            }

            Shape newEquation = null;
            if (latexCode.Trim() != "") {
                newEquation = GetPictureShapeFromLaTeXCode( slide, latexCode, editor.FontSize );
            }

            if (newEquation != null) {
                LaTeXTags tags = newEquation.LaTeXTags();

                tags.OriginalWidth.value = newEquation.Width;
                tags.OriginalHeight.value = newEquation.Height;
                tags.FontSize.value = editor.FontSize;

                tags.Type.value = EquationType.Equation;
            }
            else {
                newEquation = CreateEmptyEquation();
            }

            newEquation.LaTeXTags().Code.value = latexCode;

            newEquation.Top = equation.Top;
            newEquation.Left = equation.Left;

            // keep the equation's scale
            // TODO: this scales everything twice if we are not careful [3/4/2009 Andreas]
            float widthScale = equation.Width / equation.LaTeXTags().OriginalWidth;
            float heightScale = equation.Height / equation.LaTeXTags().OriginalHeight;
            newEquation.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            newEquation.Width *= widthScale;
            newEquation.Height *= heightScale;

            // copy animations over from the old equation
            Sequence sequence = slide.TimeLine.MainSequence;
            var effects =
                from Effect effect in sequence
                where effect.Shape == equation
                select effect;

            CopyEffectsTo(newEquation, false, sequence, effects);

            // delete the old equation
            equation.Delete();

            // return the new equation
            return newEquation;
        }

        public List<Shape> GetInlineShapes(Shape shape)
        {
            List<Shape> shapes = new List<Shape>();

            Slide slide = shape.GetSlide();
            if (slide == null)
            {
                return shapes;
            }

            foreach (LaTeXEntry entry in shape.LaTeXTags().Entries)
            {
                if (!IsEscapeCode(entry.Code))
                {
                    Shape inlineShape = slide.Shapes.FindById(entry.ShapeId);
                    Debug.Assert(inlineShape != null);
                    if (inlineShape != null)
                    {
                        shapes.Add(inlineShape);
                    }
                }
            }
            return shapes;
        }

        /// <summary>
        /// get the shape specified by the LinkID field in LaTeXTags()
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public Shape GetLinkShape(Shape shape)
        {
            LaTeXTags tags = shape.LaTeXTags();
            // pre-condition: shape has a linked shape
            Debug.Assert(tags.Type == EquationType.Inline || tags.Type == EquationType.Equation);
            Slide slide = shape.GetSlide();
            Trace.Assert(slide != null);

            Shape linkShape = slide.Shapes.FindById(tags.LinkID);
            Trace.Assert(linkShape != null);
            return linkShape;
        }

        private ShapeRange GetShapeRange<T>(T list) where T : IEnumerable<Shape>
        {
            Selection selection = Application.ActiveWindow.Selection;
            selection.Unselect();

            foreach (Shape shape in list)
            {
                shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            }

            // FIXME: ignore ChildShapeRange for now [12/31/2008 Andreas]
            ShapeRange shapeRange = selection.ShapeRange;
            selection.Unselect();

            return shapeRange;
        }
    }
}
