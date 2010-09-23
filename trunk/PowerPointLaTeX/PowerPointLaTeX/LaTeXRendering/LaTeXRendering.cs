using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class LaTeXRendering
    {
        // renders all formulas at a higher resolution than necessary to allow for zooming
        public const int RenderDPISetting = 300;

        static public Shape GetPictureShapeFromLaTeXCode(Slide currentSlide, string latexCode, float fontSize)
        {
            // check the cache first
            int baselineOffset;
            float wantedPixelsPerEmHeight = DPIHelper.FontSizeToPixelsPerEmHeight(fontSize, DPIHelper.WindowsDPISetting);
            float actualPixelsPerEmHeight = DPIHelper.FontSizeToPixelsPerEmHeight(fontSize, RenderDPISetting);
            Image image = GetImageForLaTeXCode(latexCode, ref actualPixelsPerEmHeight, out baselineOffset);
            if (image == null)
            {
                return null;
            }

            Shape picture = CreatePictureShapeFromImage(currentSlide, image);
            if (picture == null)
            {
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
            picture.LaTeXTags().BaseLineOffset.value = (float)baselineOffset / image.Height;
            //picture.LaTeXTags().PixelsPerEmHeight = actualPixelsPerEmHeight;
            string shortenedName;
            if (latexCode.Length > 32)
            {
                shortenedName = latexCode.Substring(0, 32) + "..";
            }
            else
            {
                shortenedName = latexCode;
            }
            picture.Name = "LaTeX: " + shortenedName;

            return picture;
        }

        /// <returns>A valid picture shape from the data or null if creation failed</returns>
        static private Shape CreatePictureShapeFromImage(Slide slide, Image image)
        {
            IDataObject oldClipboardContent = null;
            try
            {
                oldClipboardContent = Clipboard.GetDataObject();
            }
            catch { Debug.Assert(false, "Retrieving the current clipboard contents failed!"); }

            Clipboard.SetImage(image);

            ShapeRange pictureRange = slide.Shapes.Paste();
            if (oldClipboardContent != null)
                Clipboard.SetDataObject(oldClipboardContent);

            if (pictureRange == null)
            {
                return null;
            }

            Trace.Assert(pictureRange.Count == 1);
            return pictureRange[1];
        }

        static public Image GetImageForLaTeXCode(string latexCode, ref float pixelsPerEmHeight, out int baselineOffset)
        {
            byte[] imageData = GetImageDataForLaTeXCode(latexCode, ref pixelsPerEmHeight, out baselineOffset);
            return GetImageFromImageData(imageData);
        }

        /// <summary>
        /// Get the raw image data for some latex code or null if the compilation failed.
        /// It seems to be possible to receive a byte[0] array for some reason.
        /// </summary>
        /// <param name="latexCode"></param>
        /// <returns></returns>
        static private byte[] GetImageDataForLaTeXCode(string latexCode, ref float pixelsPerEmHeight, out int baselineOffset)
        {
            // TODO: this is very much a hack! (to allow everything to stay static) [9/22/2010 Andreas]
            CacheTags presentationCache = Globals.ThisAddIn.Tool.ActivePresentation.CacheTags();

            byte[] imageData;

            byte[] cachedImageData = null;
            float cachedPixelsPerEmHeight = 0;
            int cachedBaselineOffset = 0;
            // TODO: rewrite the cache system to work even if the main thread is blocked [8/4/2009 Andreas]
            if (presentationCache[latexCode].IsCached())
            {
                presentationCache[latexCode].Query(out cachedImageData, out cachedPixelsPerEmHeight, out cachedBaselineOffset);

                // make sure we return a some-what meaningful array
                if (cachedImageData != null && cachedImageData.Length == 0)
                {
                    cachedImageData = null;
                }
            }

            // convert to int to avoid floating point issues [5/24/2010 Andreas]
            if (cachedImageData != null && (int)cachedPixelsPerEmHeight >= (int)pixelsPerEmHeight)
            {
                // we can use the cached formula
                imageData = cachedImageData;
                pixelsPerEmHeight = cachedPixelsPerEmHeight;
                baselineOffset = cachedBaselineOffset;

                presentationCache[latexCode].Use();
            }
            else
            {
                // try to render the formula using our LaTeX service
                Globals.ThisAddIn.LaTeXRenderingServices.Service.RenderLaTeXCode(latexCode, out imageData, ref pixelsPerEmHeight, out baselineOffset);

                if (imageData != null && imageData.Length > 0)
                {
                    // looks good, so cache it
                    presentationCache[latexCode].Store(imageData, pixelsPerEmHeight, baselineOffset);
                }
                else
                {
                    // if this failed, use the result from the cache, can't be off worse
                    imageData = cachedImageData;
                    pixelsPerEmHeight = cachedPixelsPerEmHeight;
                    baselineOffset = cachedBaselineOffset;

                    if (cachedImageData != null)
                    {
                        presentationCache[latexCode].Use();
                    }
                }
            }

            return imageData;
        }


        static private Image GetImageFromImageData(byte[] imageData)
        {
            if (imageData == null)
            {
                return null;
            }

            MemoryStream stream = new MemoryStream(imageData);
            Image image;
            try
            {
                image = Image.FromStream(stream);
            }
            catch
            {
                return null;
            }
            return image;
        }

    }
}
