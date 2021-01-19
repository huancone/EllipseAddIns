namespace System.Web.Services.Ellipse
{
    public class ClientConversation
    {
        public static string Username;
        public static string Password;
        public static string District;
        public static string Position;
        //TO DO New District Variable
        public static bool debuggingMode;

        public static void authenticate(string username, string password, string district, string position)
        {
            ClientConversation.Username = username;
            ClientConversation.Password = password;
            ClientConversation.District = district;
            ClientConversation.Position = position;
        }

        public static void authenticate(string username, string password)
        {
            ClientConversation.Username = username;
            ClientConversation.Password = password;
        }

        public static void StartDebugging()
        {
            ClientConversation.debuggingMode = true;
        }
        public static void StopDebugging()
        {
            ClientConversation.debuggingMode = false;
        }

    }
}
