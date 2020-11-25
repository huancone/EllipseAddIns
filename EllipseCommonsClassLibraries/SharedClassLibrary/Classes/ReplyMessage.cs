//Shared Class Library - ReplyMessage
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

using System.CodeDom;
using System.Collections.Generic;
using System.Linq;

namespace SharedClassLibrary.Classes
{
    public class ReplyMessage
    {
        public string[] Errors
        {
            get => _errorList?.ToArray();
            set
            {
                if (value == null)
                    _errorList = null;
                else
                {
                    _errorList = new List<string>();
                    foreach(var item in value)
                        _errorList.Add(item);
                }
            }
        }
        public string Message;
        public string[] Warnings
        {
            get => _warningList?.ToArray();
            set
            {
                if (value == null)
                    _warningList = null;
                else
                {
                    _warningList = new List<string>();
                    foreach (var item in value)
                        _warningList.Add(item);
                }
            }
        }
        private List<string> _errorList = new List<string>();
        private List<string> _warningList = new List<string>();

        public string GetStringErrors()
        {
            if (Errors == null)
                return null;

            var message = Errors.Aggregate("", (current, error) => current + (error + ". "));

            return message.Trim();
        }

        public string GetStringWarnings()
        {
            if (Warnings == null)
                return null;

            var message = Warnings.Aggregate("", (current, warning) => current + (warning + ". "));

            return message.Trim();
        }

        public void AddError(string errorMessage)
        {
            if(_errorList == null)
                _errorList = new List<string>();
            _errorList.Add(errorMessage);
        }
        public void AddError(string[] errorMessages)
        {
            if (_errorList == null)
                _errorList = new List<string>();

            if (errorMessages == null)
                return;

            foreach(var item in errorMessages)
                _errorList.Add(item);
        }
        public void AddError(List<string> errorMessages)
        {
            if (_errorList == null)
                _errorList = new List<string>();

            if (errorMessages == null)
                return;

            foreach (var item in errorMessages)
                _errorList.Add(item);
        }
        public void AddWarning(string warningMessage)
        {
            if (_warningList == null)
                _warningList = new List<string>();
            _warningList.Add(warningMessage);
        }

        public void AddWarning(string[] warningMessages)
        {
            if (_warningList == null)
                _warningList = new List<string>();
            
            if (warningMessages == null)
                return;

            foreach (var item in warningMessages)
                _warningList.Add(item);
        }
        public void AddWarning(List<string> warningMessages)
        {
            if (_warningList == null)
                _warningList = new List<string>();

            if (warningMessages == null)
                return;

            foreach (var item in warningMessages)
                _warningList.Add(item);
        }
    }
}