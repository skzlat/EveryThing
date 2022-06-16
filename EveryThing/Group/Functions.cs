using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EveryThing
{
    public partial class Functions
    {
        public class Head
        {
            public int Code { get; set; }
            public string Message { get; set; }
            public string Description { get; set; }
        }

        public class GroupesClass
        {
            public Head Head { get; set; }
            public string Manager { get; set; }
            public int CommonPriceGroups { get; set; }
            public IList<string> TableHead { get; set; }
            public IList<IList<string>> Table { get; set; }
        }

        public class PriceGroup
        {
            public string Code { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string DisplayString { get; set; }
        }

        public List<PriceGroup> PriceGroupesList { get; set; }
        public List<TypesGroup> TypesList { get; set; }

        
        public char TextSplitChar(string text)
        {
            char SplitChar = char.MinValue;
            // Вычленяем символы из строки настроек для отображения типов конструкций. Берём первый для упрощения. Будет необходимость - допилю
            Regex regex = new Regex(@"[^0-9a-zA-Z]+");
            MatchCollection matches = regex.Matches(text);
            if (matches.Count > 0)
                SplitChar = char.Parse(matches[0].Value);
            return SplitChar;
        }
        public const string technicalRequestsFolder = @"\\newbuffer\buffer\SAPR\TechnicalRequests\";

        public void OpenFile(string filePath)
        {
            //Открываем созданный файл
            System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
            myProcess.StartInfo.FileName = filePath;
            //myProcess.StartInfo.Verb = "Open";
            //myProcess.StartInfo.CreateNoWindow = false;
            myProcess.Start();
        }
    }
}
