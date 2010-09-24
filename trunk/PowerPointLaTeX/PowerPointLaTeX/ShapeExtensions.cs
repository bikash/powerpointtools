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
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class ShapeExtensions
    {
        internal static List<Shape> GetInlineShapes(this Shape shape)
        {
            List<Shape> shapes = new List<Shape>();

            Slide slide = shape.GetSlide();
            if (slide == null)
            {
                return shapes;
            }

            foreach (LaTeXEntry entry in shape.LaTeXTags().Entries)
            {
                if (!Helpers.IsEscapeCode(entry.Code))
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
        internal static Shape GetLinkShape(this Shape shape)
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

        internal static void AddEffects(this Shape target, IEnumerable<Effect> effects, bool setToWithPrevious, Sequence sequence)
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
    }
}
