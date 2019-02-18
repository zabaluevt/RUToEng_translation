using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RUToEng_translation
{
    class RuStrings
    {
        public Collection<PathAndValue> collection { get; set; }

        List<string> list = new List<string>();

        public RuStrings()
        {
            collection = new Collection<PathAndValue>();
        }

        public void GetRegexedStrings(string path)
        {
            string file = ReadFile(path);

            Regex regex = new Regex("\"[А-я!\\ {01234},.]+");
            var matches = regex.Matches(file);

            Regex regexRu = new Regex("[А-я]");

            foreach (var item in matches)
            {
                if (item.ToString().Length > 4 && regexRu.IsMatch(item.ToString()))
                {
                    if (!item.ToString().Remove(0, 1).Equals(collection.LastOrDefault()?.Value))
                        collection.Add(new PathAndValue(item.ToString().Remove(0, 1), path));
                }
                   
            }
        }

        public string ReadFile(string path)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                return sr.ReadToEnd();
            }
        }

        public List<string> GetAllFilesInFolder(string str)
        {
            DirectoryInfo info = new DirectoryInfo(str);
            DirectoryInfo[] dirs = info.GetDirectories();

            foreach (var dir in dirs)
            {
                GetAllFilesInFolder(dir.FullName);
            }

            foreach (var file in info.GetFiles())
            {
                if(file.FullName.EndsWith(".cs") 
                    || file.FullName.EndsWith(".xaml"))
                    list.Add(file.FullName);
            }
                
            return list;
        }
    }
}
