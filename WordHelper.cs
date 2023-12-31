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
            dynamic app = new Word.Application();
            Object file = _fileInfo.FullName;
            Object missing = Type.Missing;
            dynamic doc = app.Documents.Open(file);

            foreach (var item in items)
            {
                Word.Find find = app.Selection.Find;
                find.Text = item.Key;
                find.Replacement.Text = item.Value;
                Object wrap = Word.WdFindWrap.wdFindContinue;
                Object replace = Word.WdReplace.wdReplaceAll;
                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing, Replace: replace);
            }

            Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd HHmmss") + _fileInfo.Name);
            app.ActiveDocument.SaveAs2(newFileName);
            DataClass.FileNamePrint = newFileName;
            app.Visible = true;
            
            doc.Activate();
            int dialogResult = app.Dialogs[Word.WdWordDialog.wdDialogFilePrint].Show();
            app.ActiveDocument.Close();

            return true;
        }
    }
}
