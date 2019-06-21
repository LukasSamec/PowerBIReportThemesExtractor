using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Layout
{
    class Layout
    {
        
        public Layout(Sections[] sections)
        {
            this.Sections = sections;
        }

        Sections[] Sections { get; set; }
    }
}
