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

namespace PowerPointLaTeX
{
    class SettingsTags
    {
        public delegate void ToggleChangedEventHandler(bool enabled);
        public static event ToggleChangedEventHandler ManualPreviewChanged = null;
        public static event ToggleChangedEventHandler PresentationModeChanged = null;
        public static event ToggleChangedEventHandler ManualEquationEditingChanged = null;
        
        // "manual" instead of "automatic" to make it automatic by default :-)
        public AddInTagBool ManualPreview;
        public AddInTagBool ManualEquationEditing;
        public AddInTagBool PresentationMode;

        public SettingsTags(Presentation presentation)
        {
            Tags tags = presentation.Tags;

            ManualPreview = new AddInTagBool(tags, "ManualPreview");
            PresentationMode = new AddInTagBool(tags, "PresentationMode");
            ManualEquationEditing = new AddInTagBool(tags, "ManualEquationEditing");

            ManualPreview.ValueChanged += new ValueChangedEventHandler<bool>(AutomaticPreview_ValueChanged);
            PresentationMode.ValueChanged += new ValueChangedEventHandler<bool>(PresentationMode_ValueChanged);
            ManualEquationEditing.ValueChanged += new ValueChangedEventHandler<bool>(ManualEquationEditing_ValueChanged);
        }

        void ManualEquationEditing_ValueChanged(object sender, bool value)
        {
            ToggleChangedEventHandler handler = ManualEquationEditingChanged;
            if (handler != null)
            {
                handler(value);
            }
        }

        void PresentationMode_ValueChanged(object sender, bool value)
        {
            ToggleChangedEventHandler handler = PresentationModeChanged;
            if (handler != null)
            {
                handler(value);
            }
        }

        void AutomaticPreview_ValueChanged(object sender, bool value)
        {
            ToggleChangedEventHandler handler = ManualPreviewChanged;
            if (handler != null)
            {
                handler(value);
            }
        }

        public void Clear()
        {
            ManualPreview.Clear();
            PresentationMode.Clear();
            ManualEquationEditing.Clear();
        }
    }
}
