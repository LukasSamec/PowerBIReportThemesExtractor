using Newtonsoft.Json;
using PowerBIReportThemesExtractor.Config;
using PowerBIReportThemesExtractor.Layout;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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
        private string defaultColor;

        public ThemeExtractor(string originalFilePath, string themePath, string defaultColor)
        {
            this.originalFilePath = originalFilePath;
            this.zipFilePath = originalFilePath + ".zip";
            this.themePath = themePath;
            this.themeName = Path.GetFileNameWithoutExtension(themePath);
            this.defaultColor = defaultColor;
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

            List<VisualConfig> uniqueVisualConfigs = visualConfigs.GroupBy(p => p.VisualTypeName).Select(g => g.First()).ToList();

            if (visualConfigs.Count != uniqueVisualConfigs.Count)
            {
                MessageBox.Show("More visuals of the same type has been found. The program will extract one visual of every type in report.", "More visuals of the same type", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            string text = this.GetThemeText(uniqueVisualConfigs);

            File.WriteAllText(themePath, text);

            MessageBox.Show("Theme has been successfully extracted to: " + themePath,"Success",MessageBoxButton.OK,MessageBoxImage.Information);
            
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
                    string propertyValue = this.ExtractVisualObjectsProperyValue(document, objectName, propertyName, visualTypeName);
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

        private string ExtractVisualObjectsProperyValue(XmlDocument document, string objectName, string objectProperty, string visualTypeName)
        {
            XmlNode node = null;

            if (objectProperty.ToLower().Contains("color"))
            {
                node = document.SelectSingleNode("/Root/singleVisual/objects/" + objectName + "/properties/" + objectProperty + "/solid/color/expr/ThemeDataColor/text() |" + "/Root/singleVisual/vcObjects/" + objectName + "/properties/" + objectProperty + "/solid/color/expr/ThemeDataColor/text()");

                if (node == null)
                {
                    string color = "\"" + defaultColor + "\"";
                    MessageBox.Show(string.Format("Color in {0} - {1} - {2} is not defined by hex code. The color will by set by default color value {3}", visualTypeName, objectName, objectProperty, defaultColor), "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return color;
                }
                return node.Value;

            }

            node = document.SelectSingleNode("/Root/singleVisual/objects/" + objectName + "/properties/" + objectProperty + "/expr/Literal/Value/text() |" +
                "/Root/singleVisual/vcObjects/" + objectName + "/properties/" + objectProperty + "/expr/Literal/Value/text()");

            return node.Value;
        }

        #endregion xml

        #region themeText

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

                if (i != configs.Count - 1)
                {
                    text += ",";
                }
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
                    if (config.VisualConfigObjects[i].Properies[j].PropertyName.ToLower().Contains("color"))
                    {
                        text += "\"" + config.VisualConfigObjects[i].Properies[j].PropertyName + "\":" + "{ \"solid\": { \"color\":" + config.VisualConfigObjects[i].Properies[j].Value + "} }";
                    }
                    else
                    {
                        text += "\"" + config.VisualConfigObjects[i].Properies[j].PropertyName + "\":" + config.VisualConfigObjects[i].Properies[j].Value;
                    }
                   
                    if (j != config.VisualConfigObjects[i].Properies.Count - 1 )
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

        #endregion themeText

    }
}
