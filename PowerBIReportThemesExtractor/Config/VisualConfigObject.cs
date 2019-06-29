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
        }

        public string ConfigObjectName { get; set; }
    }
}
