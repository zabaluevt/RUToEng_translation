using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RUToEng_translation
{
    class PathAndValue
    {
        public PathAndValue(string st1, string st2)
        {
            Path = st2;
            Value = st1;
        }
        public string Path { get; set; }
        public string Value { get; set; }
     
    }
}
