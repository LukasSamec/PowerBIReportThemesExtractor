using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PowerBIReportThemesExtractor.Config;
using PowerBIReportThemesExtractor.Layout;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace PowerBIReportThemesExtractor.Extractor
{
    public class ThemeExtractor
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
            ReportLayout layout = this.ExtractLayout();
            List<XmlDocument> visualConfigs = this.ExtractVisualConfigs(layout);
            VisualConfig config = this.GetVisualConfig(visualConfigs[0]);          
        }

        private ReportLayout ExtractLayout()
        {

            File.Copy(this.originalFilePath, this.zipFilePath);
            ZipArchive zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Read);
            ReportLayout layout = null;
        
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.Name == "Layout")
                {
                    using (StreamReader sr = new StreamReader(entry.Open(), Encoding.Unicode, true))
                    {
                        string jsonText = sr.ReadToEnd();
                        layout = JsonConvert.DeserializeObject<ReportLayout>(jsonText);
                    }
                                         
                    break;
                }
            }

            zip.Dispose();
            File.Delete(zipFilePath);

            return layout;
        }

        private List<XmlDocument> ExtractVisualConfigs(ReportLayout layout)
        {
            List<XmlDocument> configs = new List<XmlDocument>();

            foreach (Sections section in layout.Sections)
            {
                foreach (VisualContainers container in section.VisualContainers)
                {
                    XmlDocument doc = JsonConvert.DeserializeXmlNode(container.Config,"Root");
                    configs.Add(doc);
                }
            }

            return configs;
        }

        private string ExtractVisualTypeName(XmlDocument document)
        {
            return document.SelectSingleNode("/Root/singleVisual/visualType/text()").Value;
        }

        private List<string> ExtractVisualObjectsNames(XmlDocument document)
        {
            List<string> objectNames = new List<string>();
            XmlNodeList nodeList = document.SelectNodes("/Root/singleVisual/objects/*");

            foreach (XmlNode node in nodeList)
            {
                objectNames.Add(node.Name);
            }

            return objectNames;
        }

        private List<string> ExtractVisualObjectsProperiesNames(XmlDocument document,string objectName)
        {
            List<string> objectNames = new List<string>();
            XmlNodeList nodeList = document.SelectNodes("/Root/singleVisual/objects/"+objectName+"/properties/*");

            foreach (XmlNode node in nodeList)
            {
                objectNames.Add(node.Name);
            }

            return objectNames;
        }

        private string ExtractVisualObjectsProperyValue(XmlDocument document,string objectName, string objectProperty)
        {
            return document.SelectSingleNode("/Root/singleVisual/objects/"+ objectName + "/properties/"+ objectProperty + "/expr/Literal/Value/text()").Value;
        }

        private VisualConfig GetVisualConfig(XmlDocument document)
        {
            string visualTypeName = this.ExtractVisualTypeName(document);
            VisualConfig visualConfig = new VisualConfig(visualTypeName);

            List<string> objectNames = ExtractVisualObjectsNames(document);

            foreach (string objectName in objectNames)
            {
                VisualConfigObject configObject = new VisualConfigObject(objectName);

                List<string> propertiesNames = this.ExtractVisualObjectsProperiesNames(document, objectName);
                foreach (string propertyName in propertiesNames)
                {
                    string propertyValue = this.ExtractVisualObjectsProperyValue(document, objectName, propertyName);
                    configObject.Properies.Add(new VisualConfigObjectProperty(propertyName, propertyValue));
                }

                visualConfig.VisualConfigObjects.Add(configObject);                                         
            }

            return visualConfig;



        }


    }
}
