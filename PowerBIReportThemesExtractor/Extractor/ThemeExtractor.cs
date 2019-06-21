using Newtonsoft.Json;
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

        public ThemeExtractor(string originalFilePath)
        {
            this.originalFilePath = originalFilePath;
            this.zipFilePath = originalFilePath + ".zip";
        }

        public void Extract()
        {
            
        }

        private PowerBIReportThemesExtractor.Layout.Layout ExtractLayout()
        {
            if (File.Exists(this.zipFilePath))
            {
                return null;
            }

            File.Copy(this.originalFilePath, this.zipFilePath);
            ZipArchive zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Read);
            PowerBIReportThemesExtractor.Layout.Layout layout = null;
        
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.Name == "Layout")
                {
                    using (StreamReader sr = new StreamReader(entry.Open(), Encoding.Unicode, true))
                    {
                        string jsonText = sr.ReadToEnd();
                        layout = JsonConvert.DeserializeObject<PowerBIReportThemesExtractor.Layout.Layout>(jsonText);
                    }
                                         
                    break;
                }
            }

            zip.Dispose();
            File.Delete(zipFilePath);

            return layout;
        }

        private string ExtractConfigs(PowerBIReportThemesExtractor.Layout.Layout layout)
        {
            return null;
        }


    }
}
