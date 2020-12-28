using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LogsheetDatamodelLibrary.Configuration
{
    public class Login
    {
        public string User;
        public string Password;
        public string EncodedPassword;

        public Login()
        {}
        public Login(string user, string password)
        {
            User = user;
            Password = password;
        }
    }
}
