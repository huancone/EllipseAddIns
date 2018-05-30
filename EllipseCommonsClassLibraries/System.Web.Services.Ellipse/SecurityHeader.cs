using System.Xml;
using System.ServiceModel.Channels;

namespace System.Web.Services.Ellipse
{
    public class SecurityHeader : MessageHeader
    {
        public override string Name
        {
            get { return "Security"; }
        }

        public override string Namespace
        {
            get { return "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"; }
        }

        protected override void OnWriteHeaderContents(XmlDictionaryWriter writer, MessageVersion messageVersion)
        {
            writer.WriteStartElement("UsernameToken");
            writer.WriteElementString("Username", ClientConversation.username);
            writer.WriteElementString("Password", ClientConversation.password);
            writer.WriteEndElement();
        }
    }
}
