using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Config
{
    class VisualConfigObject
    {
        public VisualConfigObject(string configObjectName)
        {
            this.ConfigObjectName = configObjectName;
            this.Properies = new List<VisualConfigObjectProperty>();
        }

        public string ConfigObjectName { get; set; }
        public List<VisualConfigObjectProperty> Properies { get; set; }
    }
}
