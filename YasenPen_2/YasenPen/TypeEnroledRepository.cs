using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace YasenPen
{
    public static class TypeEnroledRepository
    {
        public static string GetInfo(string way)
        {
            string result;

            List<string> vsps = new List<string>();

            using (StreamReader sr = new StreamReader(way, Encoding.GetEncoding(1251)))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    vsps.Add(line.ToString());
                }
            }
            if (vsps.Count > 0)
            {
                result = vsps[0];
                return result;
            }
            else
            {
                result = "";
                return result;
            }
        }

        public static void WriteInfoInDoc(string txt,string way)
        {
            StreamWriter sw = new StreamWriter(way);

            sw.WriteLine(txt);

            sw.Close();
        }

       
    }
}
