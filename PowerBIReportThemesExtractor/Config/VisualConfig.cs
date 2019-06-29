using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Config
{
    class VisualConfig
    {
        public VisualConfig(string visualTypeName, List<VisualConfigObject> visualConfigObjects)
        {
            this.VisualTypeName = visualTypeName;
            this.VisualConfigObjects = visualConfigObjects;
        }

        public string VisualTypeName { get; set; }
        public List<VisualConfigObject> VisualConfigObjects { get; set; }
    }
}
