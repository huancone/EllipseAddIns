using System;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseCommonsClassLibrary
{
    /// <summary>
    /// Debugger para gestionar los logs de errores
    /// </summary>
    public static class Debugger
    {
        private static DebugError _lastError;

        public static void DebugScreen(Screen.ScreenSubmitRequestDTO request, Screen.ScreenDTO reply, string filename)
        {
            var requestJson = new JavaScriptSerializer().Serialize(request.screenFields);
            var replyJson = new JavaScriptSerializer().Serialize(reply.screenFields);
            const string filePath = @"C:\Ellipse\Debugger\";
            FileWriter.AppendTextToFile(requestJson, "ScreenRequest.txt", filePath);
            FileWriter.AppendTextToFile(replyJson, "ScreenReply.txt", filePath);
        }
        public static void LogError(string customDetails, string errorMessage, bool debugger = false)
        {
            try
            {
                const string errorFilePath = @"C:\Ellipse\Logs\";
                const string errorFileName = @"error.txt";

                var lastError = new DebugError
                {
                    CustomDetails = customDetails,
                    ErrorMessage = errorMessage,
                    DateTime = "" + DateTime.Now,
                    UrlLocation = errorFilePath + errorFileName
                };

                _lastError = lastError;

                if (debugger)
                    MessageBox.Show(lastError.CustomDetails+ ": " + lastError.ErrorMessage, "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                var stringError = lastError.DateTime + " - " + lastError.CustomDetails + " : " + lastError.ErrorMessage;

                FileWriter.CreateDirectory(errorFilePath);
                FileWriter.AppendTextToFile(stringError, errorFileName, errorFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se puede crear el Log de Error\n" + customDetails + ": " + ex + "\n" + errorMessage, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public static DebugError GetLastError()
        {
            return _lastError;
        }
    }

    public class DebugError
    {
        public string CustomDetails;
        public string ErrorMessage;
        public string DateTime;
        public string UrlLocation;
    }
}
