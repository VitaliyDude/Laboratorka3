using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using Word = Microsoft.Office.Interop.Word;

namespace rsad
{
    internal class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            //dynamic app = new Word.Application();
            //Word.Application app1 = null;
            //try
            //{
            //app = new Word.Application();
            dynamic app = new Word.Application();
            Object file = _fileInfo.FullName;
            Object missing = Type.Missing;
            dynamic doc = app.Documents.Open(file);
            
        }
    }
}
