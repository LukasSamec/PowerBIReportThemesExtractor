using Newtonsoft.Json;
using PowerBIReportThemesExtractor.Config;
using PowerBIReportThemesExtractor.Layout;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;

namespace PowerBIReportThemesExtractor.Extractor
{
    public class ThemeExtractor
    {
        private string originalFilePath;
        private string zipFilePath;
        private string themePath;
        private string themeName;

        public ThemeExtractor(string originalFilePath, string themePath)
        {
            this.originalFilePath = originalFilePath;
            this.zipFilePath = originalFilePath + ".zip";
            this.themePath = themePath;
            this.themeName = Path.GetFileNameWithoutExtension(themePath);
        }

        public void Extract()
        {
            ReportLayout layout = this.ExtractLayout();
            List<XmlDocument> visualConfigsXml = this.ExtractVisualConfigs(layout);
            List<VisualConfig> visualConfigs = new List<VisualConfig>();

            foreach (XmlDocument xmlDoc in visualConfigsXml)
            {
                VisualConfig config = this.GetVisualConfig(xmlDoc);
                visualConfigs.Add(config);
            }

            string text = this.GetThemeText(visualConfigs);

            File.WriteAllText(themePath, text);

            MessageBox.Show("Theme was extracted successfully.","Success",MessageBoxButton.OK,MessageBoxImage.Information);
            
        }

        #region objects

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

        #endregion objects

        #region xml

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
            XmlNodeList nodeList = document.SelectNodes("/Root/singleVisual/objects/* | /Root/singleVisual/vcObjects/*");

            foreach (XmlNode node in nodeList)
            {
                objectNames.Add(node.Name);
            }

            return objectNames;
        }

        private List<string> ExtractVisualObjectsProperiesNames(XmlDocument document,string objectName)
        {
            List<string> objectNames = new List<string>();
            XmlNodeList nodeList = document.SelectNodes
                (
                "/Root/singleVisual/objects/"+objectName+ "/properties/* |" +
                " /Root/singleVisual/vcObjects/" + objectName + "/properties/*"
                );

            foreach (XmlNode node in nodeList)
            {
                objectNames.Add(node.Name);
            }

            return objectNames;
        }

        private string ExtractVisualObjectsProperyValue(XmlDocument document,string objectName, string objectProperty)
        {
            return document.SelectSingleNode
                (
                "/Root/singleVisual/objects/"+ objectName + "/properties/"+ objectProperty + "/expr/Literal/Value/text() |" +
                "/Root/singleVisual/vcObjects/" + objectName + "/properties/" + objectProperty + "/expr/Literal/Value/text() |" +
                "/Root/singleVisual/objects/" + objectName + "/properties/" + objectProperty + "/solid/color/expr/Literal/Value/text() |" +
                "/Root/singleVisual/vcObjects/" + objectName + "/properties/" + objectProperty + "/solid/color/expr/Literal/Value/text()"
                ).Value;
        }

        #endregion xml

        private string GetThemeText(List<VisualConfig> configs)
        {
            string text = "";

            text += "{";
            text += "\"name\": \"" + themeName + "\",";
            text += "\"visualStyles\":";
            text += "{";
            for (int i = 0; i < configs.Count; i++)
            {
                text += this.GetConfigObjectText(configs[i]);
            }


            text += "}";
            text += "}";


            return text;    
        }

        private string GetConfigObjectText(VisualConfig config)
        {
            string text = "";
            text += "\""+config.VisualTypeName+ "\":";
            text += "{";
            text += "\"*\":";
            text += "{";

            for (int i = 0; i < config.VisualConfigObjects.Count; i++)
            {
                text += "\""+config.VisualConfigObjects[i].ConfigObjectName+"\":";
                text += "[{";

                for (int j = 0; j < config.VisualConfigObjects[i].Properies.Count; j++)
                {
                    text += "\"" + config.VisualConfigObjects[i].Properies[j].PropertyName + "\":" + config.VisualConfigObjects[i].Properies[j].Value;
                    if (j != config.VisualConfigObjects[i].Properies.Count -1 )
                    {
                        text += ",";
                    }
                   
                }

                text += "}]";

                if (i != config.VisualConfigObjects.Count - 1)
                {
                    text += ",";
                }
                
            }

            text += "}";
            text += "}";

           

            return text;
        }

       
    
    }
}
