using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace SharedClassLibrary.Configuration
{

    [Serializable]
    public class Options : IOptions
    {
        public List<ConfigValuePair<string, string>> OptionsList { get; set; }
        private IOptions _defaultOptions;
        public IOptions DefaultOptions
        {
            get => _defaultOptions;
            set => SetDefaultOptions(value);
        }

        private void SetDefaultOptions(IOptions defaultOptions)
        {
            _defaultOptions = defaultOptions;
        }
        public void SetOption(string key, string value)
        {
            if (OptionsList == null)
                OptionsList = new List<ConfigValuePair<string, string>>();


            //Valido primeramente si existe y lo reemplazo si existe
            foreach (var item in OptionsList)
            {
                if (item.Key != null && item.Key.Equals(key))
                {
                    var newItem = new ConfigValuePair<string, string>(item.Key, value);
                    OptionsList[OptionsList.IndexOf(item)] = newItem;
                    return;
                }
            }

            //Lo agrego si no existe
            OptionsList.Add(new ConfigValuePair<string, string>(key, value));
        }


        public string GetOptionValue(ConfigValuePair<string, string> defaultOptionItem)
        {
            if (OptionsList == null)
                OptionsList = new List<ConfigValuePair<string, string>>();

            var currentItemValue = GetOptionValue(defaultOptionItem.Key);
            if (!string.IsNullOrWhiteSpace(currentItemValue))
                return currentItemValue;

            SetOption(defaultOptionItem.Key, defaultOptionItem.Value);
            return defaultOptionItem.Value;
        }

        public string GetOptionValue(string key)
        {
            foreach (var item in OptionsList)
                if (item.Key.Equals(key))
                    return item.Value;

            if (_defaultOptions != null)
            {
                var defaultItem = _defaultOptions.GetOption(key);
                if (defaultItem.Key != null)
                {
                    SetOption(defaultItem.Key, defaultItem.Value);
                    return defaultItem.Value;
                }
            }

            return null;
        }

        public ConfigValuePair<string, string> GetOption(string key)
        {
            foreach (var item in OptionsList)
                if (item.Key.Equals(key))
                    return item;
            if (_defaultOptions != null)
            {
                var defaultItem = _defaultOptions.GetOption(key);
                if (defaultItem.Key != null)
                {
                    SetOption(defaultItem.Key, defaultItem.Value);
                    return defaultItem;
                }
            }

            return new ConfigValuePair<string, string>();
        }
    }

    [Serializable]
    [XmlType(TypeName = "optionItem")]
    public struct ConfigValuePair<TK, TV>
    {
        public ConfigValuePair(TK key, TV value)
        {
            Key = key;
            Value = value;
        }
        public TK Key
        { get; set; }

        public TV Value
        { get; set; }
    }

}
