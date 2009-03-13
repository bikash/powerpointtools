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
