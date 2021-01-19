using System.IO;
using System.Web.Services.Protocols;
using System.Xml;

namespace System.Web.Services.Ellipse
{
    public class AuthenticationExtension : SoapExtension
    {
        private Stream _inwardStream;
        private Stream _outwardStream;
        private static bool _debugginMode = false;
        private static string _logPath = @"c:\ellipse\logs";
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
            this._outwardStream = stream;
            this._inwardStream = (Stream) new MemoryStream();
            return this._inwardStream;
        }

        public override void ProcessMessage(SoapMessage message)
        {
            if (!(message is SoapClientMessage))
                return;
            _debugginMode = ClientConversation.debuggingMode;

            switch (message.Stage)
            {
                case SoapMessageStage.BeforeSerialize:
                    if (_debugginMode)
                        Log(message, "BeforeSerialize");
                    break;
                case SoapMessageStage.AfterSerialize:
                    this.AfterSerialize();
                    if (_debugginMode)
                        Log(message, "AfterSerialize");
                    break;
                case SoapMessageStage.BeforeDeserialize:
                    this.BeforeDeserialize();
                    if (_debugginMode)
                        Log(message, "BeforeDeserialize");
                    break;
                case SoapMessageStage.AfterDeserialize:
                    if (_debugginMode)
                        Log(message, "AfterDeserialize");
                    break;
                default:
                    //do nothing
                    break;
            }

            
        }

        private void BeforeDeserialize()
        {
            var streamReader = new StreamReader(this._outwardStream);
            var streamWriter = new StreamWriter(this._inwardStream);
            var str = streamReader.ReadToEnd();
            streamWriter.Write(str);
            streamWriter.Flush();
            this._inwardStream.Position = 0L;
        }

        private void AfterSerialize()
        {
            var xDoc = new XmlDocument();
            this._inwardStream.Position = 0L;
            var streamReader = new StreamReader(this._inwardStream);
            var streamWriter = new StreamWriter(this._outwardStream);
            xDoc.Load((TextReader) streamReader);
            var xmlDocumentEx = new XmlDocumentEx(xDoc);
            const string str = "http://schemas.xmlsoap.org/soap/envelope/";
            const string uri = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";
            var elementsByTagName = xDoc.GetElementsByTagName("Envelope", str);
            var xmlNodeEx = xmlDocumentEx.AppendChild(elementsByTagName[0], new QName("Header", str)).AppendChild(new QName("Security", uri)).AppendChild(new QName("UsernameToken", uri));
            xmlNodeEx.SetNode(new QName("Username", uri), ClientConversation.Username);
            xmlNodeEx.SetNode(new QName("Password", uri), ClientConversation.Password);
            var innerXml = xDoc.InnerXml;
            streamWriter.Write(innerXml);
            streamWriter.Flush();
        }

        #region debuggingMethods
        private static void LogDebugging(string content)
        {
            try
            {
                if (!_debugginMode)
                    return;

                var debugFilePath = _logPath;
                var debugFileName = @"debug" + System.DateTime.Today.ToString("yyyyMMdd") + ".txt";

                var dateTime = System.DateTime.Now.ToString("yyyyMMdd hhmmss");

                var stringContent = dateTime + "  : " + content;

                FileWriter.CreateDirectory(debugFilePath);
                FileWriter.AppendTextToFile(stringContent, debugFileName, debugFilePath);
            }
            catch
            {
                // ignored
            }
        }

        private void Log(SoapMessage message, string stage)
        {

            _inwardStream.Position = 0;
            string contents = (message is SoapServerMessage) ? "SoapRequest " : "SoapResponse ";
            contents += stage + ";";

            var reader = new StreamReader(_inwardStream);

            contents += reader.ReadToEnd();

            _inwardStream.Position = 0;

            //log.Debug(contents);
            LogDebugging("url:" + message.Url);
            LogDebugging(contents);
        }

        #endregion debuggingMethods
    }

    #region xmlHandlers
    public class QName
    {
        public string Name;
        public string Uri;

        public QName(string name, string uri)
        {
            this.Name = name;
            this.Uri = uri;
        }
    }
    public class XmlNodeEx
    {
        private XmlDocumentEx Doc;
        private XmlElement Node;

        public XmlNodeEx(XmlDocumentEx doc, XmlElement node)
        {
            this.Doc = doc;
            this.Node = node;
        }

        public XmlNodeEx AppendChild(QName qname)
        {
            return this.Doc.AppendChild((XmlNode)this.Node, qname);
        }

        public void SetNode(QName qname, string value)
        {
            this.AppendChild(qname).Node.InnerText = value;
        }

        public void SetAttribute(string name, string value)
        {
            this.Node.SetAttribute(name, value);
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
            XmlElement element = this.xDoc.CreateElement(qname.Name, qname.Uri);
            parent.InsertBefore((XmlNode)element, parent.FirstChild);
            return new XmlNodeEx(this, element);
        }
    }
    #endregion
}