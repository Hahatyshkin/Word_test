﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Word_test
{
    internal class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string filename)
        {
            if (File.Exists(filename))
            {
                _fileInfo = new FileInfo(filename);
            }
            else
            {
                throw new Exception("File not found");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);

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
                        MatchSoundsLike: false,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);


                }


                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd HHmmss ") + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();


                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }

            }
            return false;
        }

        internal List<string> ReadText()
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                Word.Document doc = app.Documents.Open(file);
                var list = new List<string>();
                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    list.Add(paragraph.Range.Text);
                }

                app.ActiveDocument.Close();


                return list;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }

            }
            return null;
        }
    }
}
