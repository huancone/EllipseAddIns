using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CommonsClassLibrary.Utilities
{
    public class MyKeyValuePair<TKey, TVal>
    {
        public TKey Id { get; set; }
        public TVal Value { get; set; }

        public MyKeyValuePair() { }

        public MyKeyValuePair(TKey key, TVal val)
        {
            this.Id = key;
            this.Value = val;
        }
    }
}
