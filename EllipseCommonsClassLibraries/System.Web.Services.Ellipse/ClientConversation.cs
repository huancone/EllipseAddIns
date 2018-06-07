namespace System.Web.Services.Ellipse
{
    public class ClientConversation
    {
        public static string username;
        public static string password;
        public static string district;
        public static string position;
        //TO DO New District Variable

        public static void authenticate(string username, string password, string district, string position)
        {
            ClientConversation.username = username;
            ClientConversation.password = password;
            ClientConversation.district = district;
            ClientConversation.position = position;
        }

        public static void authenticate(string username, string password)
        {
            ClientConversation.username = username;
            ClientConversation.password = password;
        }
    }
}
