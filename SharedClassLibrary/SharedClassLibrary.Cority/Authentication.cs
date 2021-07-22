namespace SharedClassLibrary.Cority
{
    public class Authentication
    {
        internal static string _username;
        internal static string _password;
        public static void Authenticate(string username, string password)
        {
            _username = username;
            _password = password;
        }
    }
}
