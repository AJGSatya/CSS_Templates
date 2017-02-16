using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTemplates.TemplateSyncTool.Settings
{
    public class Fragment
    {
        public FragmentType Type { get; set; }
        public string Source { get; set; }
        public string Destination { get; set; }
    }
}
