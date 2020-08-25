using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace System.Web.Services.Ellipse.Post
{
    public class Message
    {
        public Message(string field, string index, string stackTrace, string text)
        {
            Field = field;
            Index = index;
            StackTrace = stackTrace;
            Text = text;
        }

        public string Field { get; set; }
        public string Index { get; set; }
        public string StackTrace { get; set; }
        public string Text { get; set; }

        public static List<Message> GetMessagesByXElements(IEnumerable<XElement> elements)
        {
            var messages = new List<Message>();
            foreach (var element in elements)
            {
                try
                {
                    var messageElement = element.Element(XName.Get("message"));
                    if (messageElement == null) continue;
                    var text = messageElement.Element(XName.Get("text"))?.Value;
                    var stacktrace = messageElement.Element(XName.Get("stacktrace"))?.Value;
                    var field = messageElement.Element(XName.Get("field"))?.Value;
                    var index = messageElement.Element(XName.Get("index"))?.Value;

                    messages.Add(new Message(field, index, stacktrace, text));
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            return messages;
        }
    }
}
