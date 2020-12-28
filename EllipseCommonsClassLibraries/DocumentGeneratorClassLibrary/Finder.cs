using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using SharedClassLibrary.Utilities;
using Microsoft.Office.Interop.Word;

namespace DocumentGeneratorClassLibrary
{
    public static class Finder
    {
        //   {{$x}}	    DirectReplace
        //   {{$x:y}}	DirectOptionReplacement
        //   {<$x/>}	SentenceReplace
        //   {<$x:y/>}	SentenceOptionReplacement
        //   {<$x>}	    SectionReplacement {</$x>}
        //   {<$x:y>}	SectionOptionReplacement {</$x:y>}

        private static void FindAndReplaceSelection(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            //execute find and replace
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Execute(ref findText, findParams.MatchCase, findParams.MatchWholeWord,
                findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, ref replaceWithText, findParams.Replace,
                findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
        }

        public static bool FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            return FindAndReplace(wordApp, findText, replaceWithText, findParams);
        }
        public static bool FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText, FindParameters findParams)
        {
            //execute find and replace
            wordApp.ActiveDocument.Range().Find.ClearFormatting();
            return wordApp.ActiveDocument.Range().Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, replaceWithText, findParams.Replace,
                findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
        }


        public static void Execute(string urlPath, string fileName, string destPath, string destFileName, List<KeyValuePair<string, string>> variableList, string docType = "PDF")
        {
            var destPathNormalized = FileWriter.NormalizePath(destPath);

            FileWriter.CreateDirectory(destPathNormalized);
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(Path.Combine(urlPath, fileName), ReadOnly: false, Visible: true);
            try
            {
                //aDoc.Activate();

                foreach (var item in variableList)
                {
                    //HeaderFooterFindAndReplace(wordApp, "{{$" + item.Key + "}}", item.Value);
                    if (item.Key.Equals("objetoContrato"))
                        FindAndReplace(wordApp, "{{$" + item.Key + "}}", item.Value, WdUnits.wdLine);
                    else
                        FindAndReplace(wordApp, "{{$" + item.Key + "}}", item.Value);
                }

                var fileType = WdSaveFormat.wdFormatDocumentDefault;
                if (docType.Equals("PDF"))
                    fileType = WdSaveFormat.wdFormatPDF;

                SaveAs(aDoc, Path.Combine(destPathNormalized, destFileName), fileType, new SaveAsParameters());
                wordApp.Quit();
            }
            catch
            {
                aDoc.Close();
                wordApp.Quit();
                throw;
            }
        }


        public static void SaveAs(Microsoft.Office.Interop.Word.Document doc, string urlFileName, WdSaveFormat fileFormat, SaveAsParameters saveAsParameters)
        {
            doc.SaveAs(urlFileName, fileFormat, saveAsParameters.LockComments, saveAsParameters.Password, saveAsParameters.AddToRecentFilesfalse, saveAsParameters.WritePassword, saveAsParameters.ReadOnlyRecommended, saveAsParameters.EmbedTrueTypeFonts, saveAsParameters.SaveNativePictureFormat, saveAsParameters.SaveFormsData, saveAsParameters.SaveAsAoceLetter, saveAsParameters.Encoding, saveAsParameters.InsertLineBreaks, saveAsParameters.AllowSubstitutions, saveAsParameters.LineEnding, saveAsParameters.AddBiDiMaks);

            doc.Close();
        }

        #region FindExpandAndReplace
        public static bool FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText, WdUnits wideUnits)
        {
            var findParams = new FindParameters();
            return FindExpandAndReplace(wordApp, findText, replaceWithText, wideUnits, findParams);
        }
        public static bool FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText, WdUnits wideUnits, FindParameters findParams)
        {
            return FindExpandAndReplace(wordApp, findText, replaceWithText, wideUnits, findParams);
        }
        public static bool FindExpandAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText, WdUnits wideUnits, FindParameters findParams)
        {
            if (wideUnits == WdUnits.wdLine)
                return FindExpandAndReplaceSelection(wordApp, findText, replaceWithText, wideUnits, findParams);

            wordApp.ActiveDocument.Range().Find.ClearFormatting();
            var r = wordApp.ActiveDocument.Range();

            var boolValue = false;
            var itemFound = true;
            while (itemFound)
            {
                itemFound = r.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                    findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, Missing.Value, WdReplace.wdReplaceNone,
                    findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
                if (!itemFound)
                    break;
                r.Expand(wideUnits); // or change to .wdSentence or .wdParagraph
                r.Text = replaceWithText.ToString();
                boolValue = true;
            }

            return boolValue;
        }

        private static bool FindExpandAndReplaceSelection(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText, WdUnits wideUnits, FindParameters findParams)
        {
            wordApp.Selection.Find.ClearFormatting();

            var s = wordApp.Selection;

            var boolValue = false;
            var itemFound = true;

            while (itemFound)
            {
                itemFound = s.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                    findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, Missing.Value, WdReplace.wdReplaceNone,
                    findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);

                if (!itemFound)
                    break;
                s.Expand(wideUnits); // or change to .wdSentence or .wdParagraph
                s.Text = replaceWithText.ToString();
                boolValue = true;
            }

            return boolValue;
        }
        #endregion

        #region FindAndDelete
        public static int FindAndDelete(Microsoft.Office.Interop.Word.Application wordApp, object findText)
        {
            var findParams = new FindParameters();
            const WdUnits wideUnits = WdUnits.wdWord;
            return FindAndDelete(wordApp, findText, wideUnits, findParams);
        }
        public static int FindAndDelete(Microsoft.Office.Interop.Word.Application wordApp, object findText, WdUnits wideUnits)
        {
            var findParams = new FindParameters();
            return FindAndDelete(wordApp, findText, wideUnits, findParams);
        }
        public static int FindAndDelete(Microsoft.Office.Interop.Word.Application wordApp, object findText, FindParameters findParams)
        {
            const WdUnits wideUnits = WdUnits.wdWord;
            return FindAndDelete(wordApp, findText, wideUnits, findParams);
        }

        public static int FindAndDelete(Microsoft.Office.Interop.Word.Application wordApp, object findText, WdUnits wideUnits, FindParameters findParams)
        {
            if (wideUnits == WdUnits.wdLine)
                return FindAndDeleteSelection(wordApp, findText, wideUnits, findParams);

            wordApp.ActiveDocument.Range().Find.ClearFormatting();
            var r = wordApp.ActiveDocument.Range();
            r.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, Missing.Value, findParams.Replace,
                findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);

            r.Expand(wideUnits); // or change to .wdSentence or .wdParagraph
            return r.Delete();
        }

        private static int FindAndDeleteSelection(Microsoft.Office.Interop.Word.Application wordApp, object findText, WdUnits wideUnits, FindParameters findParams)
        {
            wordApp.Selection.Find.ClearFormatting();

            var s = wordApp.Selection;
            s.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, Missing.Value, findParams.Replace,
                findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
            s.Expand(WdUnits.wdLine); // or change to .wdSentence or .wdParagraph
            return s.Delete();
        }
        #endregion

        #region HeaderFooterFindAndReplace
        public static void HeaderFooterFindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            //execute find and replace
            var doc = wordApp.ActiveDocument;

            foreach (Section sec in doc.Sections)
            {
                foreach (HeaderFooter item in sec.Headers)
                {
                    item.Range.Find.Execute(ref findText, findParams.MatchCase, findParams.MatchWholeWord,
                        findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, replaceWithText, findParams.Replace,
                        findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
                }

                foreach (HeaderFooter item in sec.Footers)
                {
                    item.Range.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                        findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, replaceWithText, findParams.Replace,
                        findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
                }
            }
        }

        public static void HeaderFindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            //execute find and replace
            var doc = wordApp.ActiveDocument;

            foreach (Section sec in doc.Sections)
            {
                foreach (HeaderFooter item in sec.Headers)
                {
                    item.Range.Find.Execute(ref findText, findParams.MatchCase, findParams.MatchWholeWord,
                        findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, replaceWithText, findParams.Replace,
                        findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
                }
            }
        }

        public static void FooterFindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            //execute find and replace
            var doc = wordApp.ActiveDocument;

            foreach (Section sec in doc.Sections)
            {
                foreach (HeaderFooter item in sec.Footers)
                {
                    item.Range.Find.Execute(findText, findParams.MatchCase, findParams.MatchWholeWord,
                        findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, replaceWithText, findParams.Replace,
                        findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);
                }
            }
        }
        #endregion
    }
}
