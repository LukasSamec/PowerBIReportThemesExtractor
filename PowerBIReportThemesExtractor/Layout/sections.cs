using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Layout
{
    class Sections
    {

        public Sections(VisualContainers[] visualContainers)
        {
            this.VisualContainers = visualContainers;
        }

        VisualContainers[] VisualContainers { get; set; }
    }
}
