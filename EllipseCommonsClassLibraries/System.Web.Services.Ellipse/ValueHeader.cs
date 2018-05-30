using System.Xml;
using System.ServiceModel.Channels;

namespace System.Web.Services.Ellipse
{
    public class ValueHeader : MessageHeader
    {
        private string name;
        private string value;

        public ValueHeader(string name, string value)
        {
            this.value = value;
            this.name = name;
        }

        public override string Name
        {
            get { return name; }
        }

        public override string Namespace
        {
            get { return "http://connectivity.ews.mincom.com/"; }
        }

        protected override void OnWriteHeaderContents(XmlDictionaryWriter writer, MessageVersion messageVersion)
        {
            writer.WriteAttributeString("value", value);
        }
    }
}
