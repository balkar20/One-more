using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace YasenPen
{
    public static class ResponseRepository
    {
        public static List<string> itemsForVsp = new List<string>();



        public static List<string> GetResponse()
        {
            itemsForVsp = GetfromDoc(itemsForVsp, "response.txt");


            return itemsForVsp;
        }

        private static List<string> GetfromDoc(List<string> list, string way)
        {
            using (StreamReader sr = new StreamReader(way, Encoding.GetEncoding(1251)))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    itemsForVsp.Add(line.ToString());
                }
            }

            return list;

        }

        public static void WriteResponsivesInDoc(string txt, string way)
        {
            List<string> responcieves = new List<string>();

            using (StreamReader sr = new StreamReader(way, Encoding.GetEncoding(1251)))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    responcieves.Add(line.ToString());
                }
            }

            StreamWriter writer = new StreamWriter(@way, true);

            if (responcieves.Contains(txt) == false)
                writer.WriteLine(txt);

            writer.Close();
        }
        public static List<string> SaveResponsieves(ref List<string> items, string txt)
        {
            foreach (string item in items)
            {
                if (item != txt)
                {
                    items.Add(item);
                }
            }
            return items;
        }
    }
}
