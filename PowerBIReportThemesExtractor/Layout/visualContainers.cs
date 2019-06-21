using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Layout
{
    class VisualContainers
    {
        
        public VisualContainers(string config)
        {
            this.Config = config;
        }

        public string Config { get; set; }
    }
}
