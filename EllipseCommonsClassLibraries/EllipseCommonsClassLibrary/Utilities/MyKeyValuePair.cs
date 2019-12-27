using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseCommonsClassLibrary.Utilities
{
    public class MyKeyValuePair<Key, Val>
    {
        public Key Id { get; set; }
        public Val Value { get; set; }

        public MyKeyValuePair() { }

        public MyKeyValuePair(Key key, Val val)
        {
            this.Id = key;
            this.Value = val;
        }
    }
}
