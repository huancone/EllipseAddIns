using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace System.Web.Services.Ellipse.Post
{
    public class ResponseDto
    {
        public List<Message> Informations { get; set; }
        public List<Message> Errors { get; set; }
        public List<Message> LegacyErrors { get; set; }
        public List<Message> Warnings { get; set; }
        public String ResponseString { get; set; }
        public XDocument ResponseXML { get; set; }
        public bool GotErrorMessages()
        {
            if (Errors != null && Errors.Count > 0)
            {
                return true;
            }
            return false;
        }
        public string GetStringErrorMessages()
        {
            if (!GotErrorMessages()) return null;
            string messages = null;
            foreach (var msg in Errors)
                messages += msg.Field + " " + msg.Text;
            return messages;
        }
        public bool GotInformationMessages()
        {
            if (Informations != null && Informations.Count > 0)
            {
                return true;
            }
            return false;
        }
        public string GetStringInformationMessages()
        {
            if (!GotInformationMessages()) return null;
            string messages = null;
            foreach (var msg in Informations)
                messages += msg.Field + " " + msg.Text;
            return messages;
        }
        public bool GotWarningMessages()
        {
            if (Warnings != null && Warnings.Count > 0)
            {
                return true;
            }
            return false;
        }
        public string GetStringWarningMessages()
        {
            if (!GotWarningMessages()) return null;
            string messages = null;
            foreach (var msg in Warnings)
                messages += msg.Field + " " + msg.Text;
            return messages;
        }
        public bool GotLegacyErrorMessages()
        {
            if (LegacyErrors != null && LegacyErrors.Count > 0)
            {
                return true;
            }
            return false;
        }
        public string GetStringLegacyErrorMessages()
        {
            if (!GotLegacyErrorMessages()) return null;
            string messages = null;
            foreach (var msg in LegacyErrors)
                messages += msg.Field + " " + msg.Text;
            return messages;
        }
    }
}
