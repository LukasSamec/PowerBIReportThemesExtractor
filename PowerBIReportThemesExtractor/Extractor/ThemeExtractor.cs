using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIReportThemesExtractor.Extractor
{
    class ThemeExtractor
    {
        private string originalFilePath;
        private string zipFilePath;
        private string layoutFilePath;

        public ThemeExtractor(string originalFilePath)
        {
            this.originalFilePath = originalFilePath;
            this.zipFilePath = originalFilePath + ".zip";
            this.layoutFilePath = Path.Combine(this.zipFilePath, "Report");
            string text = ExtractJsonText();
        }

        public String ExtractJsonText()
        {
            File.Copy(this.originalFilePath, this.zipFilePath);
            ZipArchive zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Read);
            string jsonText = null;
        
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.Name == "Layout")
                {
                    using (StreamReader sr = new StreamReader(entry.Open(), Encoding.Unicode, true))
                    {
                        jsonText = sr.ReadToEnd();
                    }
                                         
                    break;
                }
            }
            zip.Dispose();
            File.Delete(zipFilePath);

            return jsonText;
        }


    }
}
