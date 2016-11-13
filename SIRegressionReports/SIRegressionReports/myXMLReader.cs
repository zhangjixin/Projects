using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace SIRegressionReports
{
    public class XMLData
    {
        public string testName;
        public string duration;
        public string startTime;
        public string endTime;
        public string outcom;
        public XMLData()
        {
            testName = string.Empty;
            startTime = string.Empty;
            endTime = string.Empty;
            duration = string.Empty;
            outcom = string.Empty;
        }
    }
    class myXMLReader
    {
        public static List<XMLData> xmlReader(string filePath, Boolean pendingFlag = true)
        {
            if (File.Exists(filePath))
            {
                return xmlFileReader(filePath, pendingFlag);
            }
            else if (Directory.Exists(filePath))
            {
                List<XMLData> myList = new List<XMLData>();
                foreach (var file in Directory.GetFiles(filePath, "*", SearchOption.AllDirectories))
                {
                    myList.AddRange(xmlFileReader(file));
                }
                return myList;
            }
            return null;
        }
        public static List<XMLData> xmlFileReader(string fileName, Boolean pendingFlag = true)
        {
            List<XMLData> myList = new List<XMLData>();
            XmlTextReader xmlReader = new XmlTextReader(fileName);
            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "UnitTestResult")
                {
                    var xmlData = new XMLData();
                    xmlData.testName = xmlReader.GetAttribute("testName");
                    xmlData.duration = xmlReader.GetAttribute("duration");
                    xmlData.startTime = xmlReader.GetAttribute("startTime");
                    xmlData.endTime = xmlReader.GetAttribute("endTime");
                    xmlData.outcom = xmlReader.GetAttribute("outcome");
                    if (!pendingFlag || xmlData.outcom != "Pending")
                    {
                        myList.Add(xmlData);
                    }
                }
            }
            xmlReader.Close();
            return myList;
        }
    }
}
