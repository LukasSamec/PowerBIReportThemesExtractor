using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Layout
{
    class ReportLayout
    {
        
        public ReportLayout(Sections[] sections)
        {
            this.Sections = sections;
        }

        public Sections[] Sections { get; set; }
    }
}
