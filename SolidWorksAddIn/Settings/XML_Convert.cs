using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace Solidworks_PDM_Specification
{
    /// <summary>
    /// Класс для сохранения/загрузки настроек и документов.
    /// </summary>
    internal class XML_Convert
    {
        #region Import
        
        public void Import(out Settings settings, string path)
        {
            string xml = File.ReadAllText(path);
            settings = new Settings();

            List<XElement> ElementsInXML = XDocument.Parse(xml).Element("Settings").Element("ComparsionGlobalVariable").Descendants("KeyVaultPair").ToList();
            (string, string) tuple = ("","");
            foreach (XElement item in ElementsInXML)
            {
                tuple = Import_Dictionary(item);
                settings.ComparsionGlobalVariable[tuple.Item1] = tuple.Item2;
            }

            settings.Vault = XDocument.Parse(xml).Element("Settings").Element("Vault").Value;
            settings.excelTemplate = XDocument.Parse(xml).Element("Settings").Attribute("ExcelTemplate").Value;
        }

        public void Import(out string pathSettings)
        {

            string path = Directory.GetCurrentDirectory() + "\\PathWithSettings.xml";
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(string));
            Stream fStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            pathSettings = xmlSerializer.Deserialize(fStream).ToString();
            fStream.Close();
        }

        public void Import(out List<Element> elements,out DrawingStamp stamp, string path)
        {
            elements = new List<Element>();
            stamp = new DrawingStamp();

            string xml = File.ReadAllText(path);
            List<XElement> ElementsInXML = XDocument.Parse(xml).Element("SpecificationFile").Element("Elements").Descendants("Element").ToList();
            stamp = Import_Stamp(XDocument.Parse(xml).Element("SpecificationFile").Element("Stamp"));
            foreach (XElement item in ElementsInXML)
                elements.Add(Import_Element(item));

            stamp.usePdmFlag = Convert.ToBoolean(XDocument.Parse(xml).Element("SpecificationFile").Attribute("UsePdmFlag").Value);
        }

        private (string, string) Import_Dictionary(XElement item)
        {
            string value = item.Attribute("Value").Value;
            string key = item.Attribute("Key").Value;
            return (key, value);
        }

        private Element Import_Element(XElement item)
        {
            Element element = new Element();
            element.Name = item.Attribute("Name").Value;
            element.Designation = item.Attribute("Designation").Value;
            element.DrawingPaperSize = item.Attribute("DrawingPaperSize").Value;
            element.Section = item.Attribute("Section").Value;
            element.Note = item.Attribute("Note").Value;
            element.Zone = item.Attribute("Zone").Value;
            element.Count = Convert.ToInt32(item.Attribute("Count").Value);

            return element;
        }
        private DrawingStamp Import_Stamp(XElement item)
        {
            DrawingStamp stamp = new DrawingStamp();
            stamp.Developer = item.Attribute("Developer").Value;
            stamp.Checker = item.Attribute("Checker").Value;
            stamp.NormativeControl = item.Attribute("NormativeControl").Value;
            stamp.Approver = item.Attribute("Approver").Value;
            stamp.Litera = item.Attribute("Litera").Value;
            stamp.InvNumbOrigin = item.Attribute("InvNumbOrigin").Value;
            stamp.ReplacedInvNumb = item.Attribute("ReplacedInvNumb").Value;
            stamp.InvNumbDupl = item.Attribute("InvNumbDupl").Value;
            stamp.ReferenceNumb = item.Attribute("ReferenceNumb").Value;
            stamp.InvNumbDupl = item.Attribute("InvNumbDupl").Value;
            stamp.ReferenceNumb = item.Attribute("ReferenceNumb").Value;
            stamp.PrimaryApplication = item.Attribute("PrimaryApplication").Value;
            stamp.FilePath = item.Attribute("FilePath").Value;
            stamp.Configuration = item.Attribute("Configuration").Value;
            stamp.DateDeveloper = Convert.ToDateTime(item.Attribute("DateDeveloper").Value);
            stamp.DateChecker = Convert.ToDateTime(item.Attribute("DateChecker").Value);
            stamp.DateNormativeControl = Convert.ToDateTime(item.Attribute("DateNormativeControl").Value);
            stamp.DateApprover = Convert.ToDateTime(item.Attribute("DateApprover").Value);

            Element element = Import_Element(item.Element("Element"));
            stamp.Name = element.Name;
            stamp.Designation = element.Designation;
            stamp.DrawingPaperSize = element.DrawingPaperSize;
            stamp.Zone = element.Zone;
            stamp.Section = element.Section;
            stamp.Note = element.Note;
            stamp.Count = element.Count;
            return stamp;
        }
        #endregion

        #region Export

        public void Export(Settings settings, string path)
        {
            XElement xmlParse = new XElement("Settings");
            XElement xmlParseVault = new XElement("Vault");
            XElement xmlParseDictionary = new XElement("ComparsionGlobalVariable");
            XAttribute xAttributeExcelTemplate = new XAttribute("ExcelTemplate", settings.excelTemplate);
            xmlParseVault.Add(settings.Vault);

            foreach (KeyValuePair<string, string> keyValuePair in settings.ComparsionGlobalVariable)
                xmlParseDictionary.Add(Export_Dictionary(keyValuePair.Key, keyValuePair.Value));
            xmlParse.Add(xmlParseVault);
            xmlParse.Add(xmlParseDictionary);
            xmlParse.Add(xAttributeExcelTemplate);
            xmlParse.Save(path);
        }

        public void Export(string pathSettings)
        {
            string path = Directory.GetCurrentDirectory() + "\\PathWithSettings.xml";
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(string));
            Stream fStream = new FileStream(path, FileMode.Create, FileAccess.Write);
            xmlSerializer.Serialize(fStream, pathSettings);
            fStream.Close();
        }

        public void Export(List<Element> elements, DrawingStamp stamp, string path)
        {
            XElement xmlParse = new XElement("SpecificationFile");
            XAttribute xAttributeUsePdmflag = new XAttribute("UsePdmFlag", stamp.usePdmFlag.ToString());

            XElement xmlListElements = new XElement("Elements");
            foreach (Element element in elements)
                xmlListElements.Add(Export_Element(element));

            XElement xmlStamp = Export_Stamp(stamp);
            xmlParse.Add(xAttributeUsePdmflag);
            xmlParse.Add(xmlStamp);
            xmlParse.Add(xmlListElements);
            if (path.Substring(path.Length - 4) != ".xml")
                path += ".xml";
            xmlParse.Save(path);
        }

        private XElement Export_Dictionary(string key, string value)
        {
            XElement parseDictionary = new XElement("KeyVaultPair");
            XAttribute xAttributeKey = new XAttribute("Key", key);
            XAttribute xAttributeValue = new XAttribute("Value", value);
            parseDictionary.Add(xAttributeKey, xAttributeValue);
            return parseDictionary;
        }

        private XElement Export_Element(Element element)
        {
            XElement parseElement = new XElement("Element");
            XAttribute xAttributeName = new XAttribute("Name", element.Name);
            XAttribute xAttributeDesignation = new XAttribute("Designation", element.Designation);
            XAttribute xAttributeDrawingPaperSize = new XAttribute("DrawingPaperSize", element.DrawingPaperSize);
            XAttribute xAttributeSection = new XAttribute("Section", element.Section);
            XAttribute xAttributeNote = new XAttribute("Note", element.Note);
            XAttribute xAttributeZone = new XAttribute("Zone", element.Zone);
            XAttribute xAttributeCount = new XAttribute("Count", element.Count.ToString());

            parseElement.Add(xAttributeName);
            parseElement.Add(xAttributeDesignation);
            parseElement.Add(xAttributeDrawingPaperSize);
            parseElement.Add(xAttributeSection);
            parseElement.Add(xAttributeNote);
            parseElement.Add(xAttributeZone);
            parseElement.Add(xAttributeCount);

            return parseElement;
        }
        
        private XElement Export_Stamp(DrawingStamp stamp)
        {
            XElement parseStamp = new XElement("Stamp");
            XElement stampElement = Export_Element(stamp);
            XAttribute xAttributeDeveloper = new XAttribute("Developer", stamp.Developer);
            XAttribute xAttributeChecker = new XAttribute("Checker", stamp.Checker);
            XAttribute xAttributeNormativeControl = new XAttribute("NormativeControl", stamp.NormativeControl);
            XAttribute xAttributeApprover = new XAttribute("Approver", stamp.Approver);
            XAttribute xAttributeLitera = new XAttribute("Litera", stamp.Litera);
            XAttribute xAttributeInvNumbOrigin = new XAttribute("InvNumbOrigin", stamp.InvNumbOrigin);
            XAttribute xAttributeReplacedInvNumb = new XAttribute("ReplacedInvNumb", stamp.ReplacedInvNumb);
            XAttribute xAttributeInvNumbDupl = new XAttribute("InvNumbDupl", stamp.InvNumbDupl);
            XAttribute xAttributeReferenceNumb = new XAttribute("ReferenceNumb", stamp.ReferenceNumb);
            XAttribute xAttributePrimaryApplication = new XAttribute("PrimaryApplication", stamp.PrimaryApplication);
            XAttribute xAttributeFilePath = new XAttribute("FilePath", stamp.FilePath);
            XAttribute xAttributeConfiguration = new XAttribute("Configuration", stamp.Configuration);
            XAttribute xAttributeDateDeveloper = new XAttribute("DateDeveloper", stamp.DateDeveloper);
            XAttribute xAttributeDateChecker = new XAttribute("DateChecker", stamp.DateChecker);
            XAttribute xAttributeDateNormativeControl = new XAttribute("DateNormativeControl", stamp.DateNormativeControl);
            XAttribute xAttributeDateApprover = new XAttribute("DateApprover", stamp.DateApprover);

            
            parseStamp.Add(xAttributeDeveloper);
            parseStamp.Add(xAttributeChecker);
            parseStamp.Add(xAttributeNormativeControl);
            parseStamp.Add(xAttributeApprover);
            parseStamp.Add(xAttributeLitera);
            parseStamp.Add(xAttributeInvNumbOrigin);
            parseStamp.Add(xAttributeReplacedInvNumb);
            parseStamp.Add(xAttributeInvNumbDupl);
            parseStamp.Add(xAttributeReferenceNumb);
            parseStamp.Add(xAttributePrimaryApplication);
            parseStamp.Add(xAttributeFilePath);
            parseStamp.Add(xAttributeDateDeveloper);
            parseStamp.Add(xAttributeDateChecker);
            parseStamp.Add(xAttributeDateNormativeControl);
            parseStamp.Add(xAttributeDateApprover);
            parseStamp.Add(xAttributeConfiguration);
            parseStamp.Add(stampElement);

            return parseStamp;
        }

        #endregion
    }
}
