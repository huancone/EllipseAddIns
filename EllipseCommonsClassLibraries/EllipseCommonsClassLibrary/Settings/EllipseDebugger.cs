using System;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using CommonsClassLibrary;

namespace EllipseCommonsClassLibrary
{
    public class Debugger : CommonsClassLibrary.Debugger
    {
        public static void DebugScreen(Screen.ScreenSubmitRequestDTO request, Screen.ScreenDTO reply, string filename)
        {
            var requestJson = new JavaScriptSerializer().Serialize(request.screenFields);
            var replyJson = new JavaScriptSerializer().Serialize(reply.screenFields);
            var filePath = Settings.CurrentSettings.LocalDataPath + @"debugger\";
            FileWriter.AppendTextToFile(requestJson, "ScreenRequest.txt", filePath);
            FileWriter.AppendTextToFile(replyJson, "ScreenReply.txt", filePath);
        }

    }
}
