using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace PowerPointLaTeX
{
    static class EquationHandling
    {
        private const int InitialEquationFontSize = 44;

        static public Shape CreateEmptyEquation(Slide slide)
        {
            const float width = 100, height = 60;

            Shape shape = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 100, 100, width, height);
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

        static public Shape EditEquation( Shape equation, out bool cancelled ) {
            EquationEditor editor = new EquationEditor( LaTeXTool.ActivePresentation.CacheTags(), equation.LaTeXTags().Code, equation.LaTeXTags().FontSize );
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
                newEquation = LaTeXRendering.GetPictureShapeFromLaTeXCode( slide, latexCode, editor.FontSize );
            }

            if (newEquation != null) {
                LaTeXTags tags = newEquation.LaTeXTags();

                tags.OriginalWidth.value = newEquation.Width;
                tags.OriginalHeight.value = newEquation.Height;
                tags.FontSize.value = editor.FontSize;

                tags.Type.value = EquationType.Equation;
            }
            else {
                newEquation = CreateEmptyEquation(slide);
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

            newEquation.AddEffects(effects, false, sequence);

            // delete the old equation
            equation.Delete();

            // return the new equation
            return newEquation;
        }
    }
}
