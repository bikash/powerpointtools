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
        public static event ToggleChangedEventHandler AnimatorModeChanged = null;

        public AddInTagBool ManualPreview;
        public AddInTagBool PresentationMode;
        public AddInTagBool AnimatorMode;

        public SettingsTags(Presentation presentation)
        {
            Tags tags = presentation.Tags;

            ManualPreview = new AddInTagBool(tags, "ManualPreview");
            PresentationMode = new AddInTagBool(tags, "PresentationMode");
            AnimatorMode = new AddInTagBool(tags, "AnimatorMode");

            ManualPreview.ValueChanged += new ValueChangedEventHandler<bool>(AutomaticPreview_ValueChanged);
            PresentationMode.ValueChanged += new ValueChangedEventHandler<bool>(PresentationMode_ValueChanged);
            AnimatorMode.ValueChanged += new ValueChangedEventHandler<bool>(AnimatorMode_ValueChanged);
        }

        void AnimatorMode_ValueChanged(object sender, bool value)
        {
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
        }
    }
}
