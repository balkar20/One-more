using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace SalaryReport
{
    class XmlParser
    {
        public string GetCurrency (XmlDocument doc)
        {
            //XDocument xdoc = XDocument.Parse(doc.ToString());
            XmlElement xRoot = doc.DocumentElement;
            var nodeList = doc.GetElementsByTagName("Currency");
            string result = null;
            foreach (XmlNode el in nodeList)
            {
                if (el.ChildNodes[3].InnerText == "Доллар США")
                {
                    result = el.ChildNodes[4].InnerText;
                    break;
                }
            }

            return result;
            ////IEnumerable<XElement> elements = xdoc.Element("DailyExRates").Elements("Currency");
            //StringReader stringReader = null;
            //XmlSerializer serializer = null;
            //XmlTextReader xmlReader = null;
            //DataStorage obj = null;
            //try
            //{
            //    stringReader = new StringReader(doc.ToString());
            //    serializer = new XmlSerializer(typeof(DataStorage));
            //    xmlReader = new XmlTextReader(stringReader);
            //    o = serializer.Deserialize(xmlReader);

            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e);
            //    throw;
            //}

            //return (DataStorage)o;

        }
    }
}
