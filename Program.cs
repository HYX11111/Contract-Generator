using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;


namespace Generator
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            System.Windows.Forms.Application.Run(new Form1());
        }

        public class WordOperate
        {
            private Microsoft.Office.Interop.Word.Application app;
            private Microsoft.Office.Interop.Word.Document doc;
            private object missing = Missing.Value;

            public void Open_Tem(string filePath)
            {
                app = new Microsoft.Office.Interop.Word.Application();
                object file = filePath;
                doc = app.Documents.Add(ref file);
                doc.Activate();
            }

            public void Open_Standard()
            {
                app = new Microsoft.Office.Interop.Word.Application();
                doc = app.Documents.Add();
                doc.Activate();
            }

            public void Replace(string strOld, string strNew)
            {
                object objReplace = WdReplace.wdReplaceAll;
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();
                app.Selection.Find.Text = strOld;
                app.Selection.Find.Replacement.Text = strNew;
                app.Selection.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref objReplace, ref missing, ref missing, ref missing, ref missing);

                StoryRanges storyRanges = doc.StoryRanges;

                foreach (Microsoft.Office.Interop.Word.Range range in storyRanges)
                {
                    range.Find.ClearFormatting();
                    range.Find.Replacement.ClearFormatting();
                    range.Find.Text = strOld;
                    range.Find.Replacement.Text = strNew;
                    range.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref objReplace, ref missing, ref missing, ref missing, ref missing);
                }

                foreach (Section section in doc.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Find.ClearFormatting();
                    headerRange.Find.Replacement.ClearFormatting();
                    headerRange.Find.Text = strOld;
                    headerRange.Find.Replacement.Text = strNew;
                    headerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref objReplace, ref missing, ref missing, ref missing, ref missing);
                    
                    Microsoft.Office.Interop.Word.Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Find.ClearFormatting();
                    footerRange.Find.Replacement.ClearFormatting();
                    footerRange.Find.Text = strOld;
                    footerRange.Find.Replacement.Text = strNew;
                    footerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref objReplace, ref missing, ref missing, ref missing, ref missing);
                }
            }

            public void Combine(string fPath0, string fPath1, string fPath2)
            {
                object pBreak = (int)WdBreakType.wdSectionBreakNextPage;

                app.Selection.InsertFile(fPath0, ref missing, false, false, false);
                app.Selection.InsertBreak(ref pBreak);
                app.Selection.InsertFile(fPath1, ref missing, false, false, false);
                app.Selection.InsertFile(fPath2, ref missing, false, false, false);
            }

            public void addPageNo()
            {
                int i = 1;
                object PageNumberAlignment = 1;
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    PageNumbers PageNo = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers;

                    if (i == 1)
                    {
                        PageNo.RestartNumberingAtSection = true;
                        PageNo.Add(ref PageNumberAlignment, false);
                    }
                    else
                    {
                        PageNo.RestartNumberingAtSection = false;
                        PageNo.Add(ref PageNumberAlignment, true);
                    }
                    i++;
                }
            }

            public void update_page()
            {
                TableOfContents table = doc.TablesOfContents[1];
                table.Update();

            }
            public void Save(string result_file)
            {
                object resultF = result_file;
                doc.SaveAs(resultF);
                Close();
                KillWinword();
            }

            public void Close()
            {
                doc.Close(ref missing, ref missing, ref missing);
                app.Quit(ref missing, ref missing, ref missing);
            }

            public void KillWinword()
            {
                Process[] p = Process.GetProcessesByName("WINWORD");
                if (p.Any())
                {
                    p[0].Kill();
                }
            }
        }

    }
}