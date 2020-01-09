using System;
using System.IO;
using System.Web.Services.Protocols;
using System.Xml;
using SoapMessage = System.Web.Services.Protocols.SoapMessage;

namespace System.Web.Services.Ellipse
{
    public class AuthenticationExtension : SoapExtension
    {
        private Stream inwardStream;
        private Stream outwardStream;

        public override void Initialize(object initializer)

        {
        }

        public override object GetInitializer(Type serviceType)
        {
            return (object) null;
        }

        public override object GetInitializer(LogicalMethodInfo methodInfo, SoapExtensionAttribute attribute)
        {
            return (object) null;
        }

        public override Stream ChainStream(Stream stream)
        { 
            this.outwardStream = stream;
            this.inwardStream = (Stream) new MemoryStream();
            return this.inwardStream;
        }

        public override void ProcessMessage(SoapMessage message)
        {
            if (!(message is SoapClientMessage))
                return;
            switch (message.Stage)
            {
                case SoapMessageStage.AfterSerialize:
                    this.afterSerialize();
                    break;
                case SoapMessageStage.BeforeDeserialize:
                    this.beforeDeserialize();
                    break;
            }
        }

        private void beforeDeserialize()
        {
            StreamReader streamReader = new StreamReader(this.outwardStream);
            StreamWriter streamWriter = new StreamWriter(this.inwardStream);
            string str = streamReader.ReadToEnd();
            streamWriter.Write(str);
            streamWriter.Flush();
            this.inwardStream.Position = 0L;
        }

        private void afterSerialize()
        {
            XmlDocument xDoc = new XmlDocument();
            this.inwardStream.Position = 0L;
            StreamReader streamReader = new StreamReader(this.inwardStream);
            StreamWriter streamWriter = new StreamWriter(this.outwardStream);
            xDoc.Load((TextReader) streamReader);
            XmlDocumentEx xmlDocumentEx = new XmlDocumentEx(xDoc);
            string str = "http://schemas.xmlsoap.org/soap/envelope/";
            string uri = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";
            XmlNodeList elementsByTagName = xDoc.GetElementsByTagName("Envelope", str);
            XmlNodeEx xmlNodeEx = xmlDocumentEx.AppendChild(elementsByTagName[0], new QName("Header", str)).AppendChild(new QName("Security", uri)).AppendChild(new QName("UsernameToken", uri));
            xmlNodeEx.SetNode(new QName("Username", uri), ClientConversation.username);
            xmlNodeEx.SetNode(new QName("Password", uri), ClientConversation.password);
            string innerXml = xDoc.InnerXml;
            streamWriter.Write(innerXml);
            streamWriter.Flush();
        }
    }
    public class QName
    {
        public string name;
        public string uri;

        public QName(string name, string uri)
        {
            this.name = name;
            this.uri = uri;
        }
    }
    public class XmlNodeEx
    {
        private XmlDocumentEx doc;
        private XmlElement node;

        public XmlNodeEx(XmlDocumentEx doc, XmlElement node)
        {
            this.doc = doc;
            this.node = node;
        }

        public XmlNodeEx AppendChild(QName qname)
        {
            return this.doc.AppendChild((XmlNode)this.node, qname);
        }

        public void SetNode(QName qname, string value)
        {
            this.AppendChild(qname).node.InnerText = value;
        }

        public void SetAttribute(string name, string value)
        {
            this.node.SetAttribute(name, value);
        }
    }
    public class XmlDocumentEx
    {
        private XmlDocument xDoc;

        public XmlDocumentEx(XmlDocument xDoc)
        {
            this.xDoc = xDoc;
        }

        public XmlNodeEx AppendChild(XmlNode parent, QName qname)
        {
            XmlElement element = this.xDoc.CreateElement(qname.name, qname.uri);
            parent.InsertBefore((XmlNode)element, parent.FirstChild);
            return new XmlNodeEx(this, element);
        }
    }
}