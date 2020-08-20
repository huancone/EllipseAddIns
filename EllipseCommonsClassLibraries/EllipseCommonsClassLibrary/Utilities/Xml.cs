using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace CommonsClassLibrary.Utilities
{
    public partial class MyUtilities
    {
        public static class Xml
        {
            /// <summary>
            /// 
            /// </summary>
            /// <param name="xDocument"></param>
            /// <param name="xPathString">Ej ("/interaction/actions/action/data/result/dto")</param>
            public static XmlNodeList GetNodeList(string xmlString, string xPathString)
            {
                XmlDocument xml = new XmlDocument();
                xml.LoadXml(xmlString);
                
                XmlNodeList xnList = xml.SelectNodes(xPathString);

                return xnList;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="xDocument"></param>
            /// <param name="xPathString">Ej ("/interaction/actions/action/data/result/dto")</param>
            /// <returns></returns>
            public static XmlNodeList GetNodeList(XDocument xDocument, string xPathString)
            {
                return GetNodeList(xDocument.ToString(), xPathString);
            }
            public static XmlNode SerializeObjectToXmlNode(Object obj)
            {
                if (obj == null)
                    throw new ArgumentNullException("Argument cannot be null");

                XmlNode resultNode = null;
                XmlSerializer xmlSerializer = new XmlSerializer(obj.GetType());
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    try
                    {
                        xmlSerializer.Serialize(memoryStream, obj);
                    }
                    catch (InvalidOperationException)
                    {
                        return null;
                    }
                    memoryStream.Position = 0;
                    XmlDocument doc = new XmlDocument();
                    doc.Load(memoryStream);
                    resultNode = doc.DocumentElement;
                }
                return resultNode;
            }
            public static void SerializeObjectToXml(string file, object obj)
            {
                System.Xml.Serialization.XmlSerializer xs
                    = new System.Xml.Serialization.XmlSerializer(obj.GetType());
                StreamWriter writer = File.CreateText(file);
                xs.Serialize(writer, obj);
                writer.Flush();
                writer.Close();
            }

            public static Object DeSerializeXmlNodeToObject(XmlNode node, Type objectType)
            {
                if (node == null)
                    throw new ArgumentNullException("Argument cannot be null");

                XmlSerializer xmlSerializer = new XmlSerializer(objectType);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    XmlDocument doc = new XmlDocument();
                    doc.AppendChild(doc.ImportNode(node, true));
                    doc.Save(memoryStream);
                    memoryStream.Position = 0;
                    XmlReader reader = XmlReader.Create(memoryStream);
                    try
                    {
                        return xmlSerializer.Deserialize(reader);
                    }
                    catch(Exception ex)
                    {
                        Debugger.LogError("MyUtilities.Xml.cs:DeSerializeXmlNodeToObject()", ex.Message);
                        return objectType.IsByRef ? null : Activator.CreateInstance(objectType);
                    }
                }
            }

            public static Object DeserializeXmlToObject(string file, Type objectType)
            {
                var xs = new XmlSerializer(objectType);
                StreamReader reader = File.OpenText(file);

                try
                {
                    return xs.Deserialize(reader);
                }
                catch (Exception ex)
                {
                    Debugger.LogError("MyUtilities.Xml.cs:DeserializeXmlToObject()", ex.Message);
                    return objectType.IsByRef ? null : Activator.CreateInstance(objectType);
                }
            }
        }
    }
}