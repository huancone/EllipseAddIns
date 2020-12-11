using System.Collections.Generic;

namespace SharedClassLibrary.Configuration
{
    public interface IOptions
    {
        List<ConfigValuePair<string, string>> OptionsList { get; set; }
        IOptions DefaultOptions { get; set;}
        ConfigValuePair<string, string> GetOption(string key);
        string GetOptionValue(ConfigValuePair<string, string> defaultOptionItem);
        string GetOptionValue(string key);
        void SetOption(string key, string value);
    }
}