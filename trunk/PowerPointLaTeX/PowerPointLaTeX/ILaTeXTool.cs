using System;
namespace PowerPointLaTeX
{
    interface ILaTeXTool
    {
        void CompilePresentation(Microsoft.Office.Interop.PowerPoint.Presentation presentation);
        void CompileShape(Microsoft.Office.Interop.PowerPoint.Slide slide, Microsoft.Office.Interop.PowerPoint.Shape shape);
        void CompileSlide(Microsoft.Office.Interop.PowerPoint.Slide slide);
        Microsoft.Office.Interop.PowerPoint.Shape CreateEmptyEquation();
        void DecompileShape(Microsoft.Office.Interop.PowerPoint.Slide slide, Microsoft.Office.Interop.PowerPoint.Shape shape);
        void DecompileSlide(Microsoft.Office.Interop.PowerPoint.Slide slide);
        Microsoft.Office.Interop.PowerPoint.Shape EditEquation(Microsoft.Office.Interop.PowerPoint.Shape equation, out bool cancelled);
        void FinalizePresentation(Microsoft.Office.Interop.PowerPoint.Presentation presentation);
    }
}
