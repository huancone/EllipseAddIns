using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedClassLibrary.Cority
{
    public class QuestionResponse
    {
        public string Code;
        public string Value;

        public QuestionResponse()
        {

        }

        public QuestionResponse(string code, string value)
        {
            Code = code;
            Value = value;
        }
    }
}
