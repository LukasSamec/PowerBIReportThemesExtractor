
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Extractor
{
    class ThemeExtractor
    {

        string originalFilePath;
        string zipFilePath;

        public ThemeExtractor(string filePath)
        {
            originalFilePath = filePath;
            zipFilePath = filePath + ".zip";

            File.Copy(originalFilePath,zipFilePath);
        }
    }
}
