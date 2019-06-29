using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Config
{
    class VisualConfigObjectProperty
    {
        public VisualConfigObjectProperty(string propertyName, string value)
        {
            this.PropertyName = propertyName;
            this.Value = value;
        }

        public string PropertyName { get; set; }

        public string Value { get; set; }
    }
}
