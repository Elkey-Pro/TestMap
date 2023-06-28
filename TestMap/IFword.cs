using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WORD = Microsoft.Office.Interop.Word;

namespace TestMap
{
    public class IFword
    {
        public WORD.Application wordApp;

        public Document document;

        public string WordReportName { get; set; }

        public string WordReportTemplate { get;set; }

        object misValue = System.Reflection.Missing.Value;

        object oFalse = false;

        public IFword(string DocLocation)
        {
            wordApp = new WORD.Application();
            document = wordApp.Documents.Open(DocLocation);
        }

        public void InsertImage(string FindText, string ImageLocation )
        {
            WORD.Range range = document.Content;
            range.Find.ClearFormatting();
            range.Find.Text = FindText;
            range.Find.MatchCase = false;
            range.Find.MatchWholeWord = true;
            bool found = range.Find.Execute();

            if (found)
            {
                WORD.Range foundRange = range.Find.Parent;

                // Get the row below the found range

                document.InlineShapes.AddPicture(ImageLocation, ref misValue, ref misValue, foundRange);
                // clear text

                ReplaceString(FindText, "");
            }
        }

        public void InsertListImage(string FindText, List<string> ListImageLocation)
        {
            WORD.Range range = document.Content;
            range.Find.ClearFormatting();
            range.Find.Text = FindText;
            range.Find.MatchCase = false;
            range.Find.MatchWholeWord = true;
            bool found = range.Find.Execute();

            if (found)
            {
                WORD.Range foundRange = range.Find.Parent;

                // Get the row below the found range

                int height = 30;
                foreach (string item in ListImageLocation)
                {

                    System.Drawing.Image img = System.Drawing.Image.FromFile(item);

                    Double ratio = img.Width / img.Height;

                    InlineShape inlineShape = document.InlineShapes.AddPicture(item, ref misValue, ref misValue, foundRange);

                    inlineShape.Height = height;

                    inlineShape.Width = (int)(height * ratio);
                }

                // clear text

                ReplaceString(FindText, "");
            }
        }



        public void ReplaceString(string FindText, string ReplaceText )
        {
            Find find = wordApp.Selection.Find;
            find.Text = FindText;
            find.Replacement.Text = ReplaceText;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            // Perform the Find and Replace
            object replace = WdReplace.wdReplaceAll;
            object missing = Type.Missing;
            find.Execute(FindText: missing, MatchCase: false, MatchWholeWord: true,
                         MatchWildcards: false, MatchSoundsLike: missing,
                         MatchAllWordForms: false, Forward: true,
                         Wrap: WdFindWrap.wdFindContinue, Format: false,
                         ReplaceWith: missing, Replace: replace);
        }

        public void ReplaceStringWithList(List<string> ListToReplace, string FindText, string LineBreak = "\r\n")
        {
            document.Activate();

            foreach (WORD.Range range in document.StoryRanges)
            {

                string joinedText = String.Join(LineBreak, ListToReplace.ToArray());

                WORD.Find find = range.Find;
                object findText = FindText;

                object replacText = joinedText;
                object replace = WORD.WdReplace.wdReplaceAll;

                if (replacText.ToString().Length > 254)
                {
                    wordApp.Application.Selection.Find.Execute(findText, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue,
                               misValue, misValue, misValue);
                    wordApp.Application.Selection.Text = (String)(replacText);
                    wordApp.Application.Selection.Collapse();
                }
                else
                {
                    find.Execute(ref findText, ref misValue, ref misValue, ref misValue, ref oFalse, ref misValue,
                        ref misValue, ref misValue, ref misValue, ref replacText,
                        ref replace, ref misValue, ref misValue, ref misValue, ref misValue);
                }


            }
        }

    }
}
