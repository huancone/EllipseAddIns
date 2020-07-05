using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace DocumentGeneratorClassLibrary
{
    public static class Finder
    {



        private static void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            var findParams = new FindParameters();
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, findParams.MatchCase, findParams.MatchWholeWord,
                findParams.MatchWildCards, findParams.MatchSoundsLike, findParams.MatchAllWordForms, findParams.Forward, findParams.Wrap, findParams.Format, ref replaceWithText, findParams.Replace,
                findParams.MatchKashida, findParams.MatchDiacritics, findParams.MatchAlefHamza, findParams.MatchControl);

            
        }

        public static void Execute(string urlPath, string fileName, string destPath, string destFileName, List<KeyValuePair<string, string>> variableList)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(Path.Combine(urlPath, fileName), ReadOnly: false, Visible: true);
            aDoc.Activate();

            foreach (var item in variableList)
            {
                FindAndReplace(wordApp, "{{$" + item.Key + "}}", item.Value);
            }

            SaveAs(aDoc, Path.Combine(destPath, destFileName), WdSaveFormat.wdFormatDocumentDefault, new SaveAsParameters());
            wordApp.Quit();
        }


        public static void SaveAs(Microsoft.Office.Interop.Word.Document doc, string urlFileName, WdSaveFormat fileFormat, SaveAsParameters saveAsParameters)
        {
            doc.SaveAs(urlFileName, fileFormat, saveAsParameters.LockComments, saveAsParameters.Password, saveAsParameters.AddToRecentFilesfalse, saveAsParameters.WritePassword, saveAsParameters.ReadOnlyRecommended, saveAsParameters.EmbedTrueTypeFonts, saveAsParameters.SaveNativePictureFormat, saveAsParameters.SaveFormsData, saveAsParameters.SaveAsAoceLetter, saveAsParameters.Encoding, saveAsParameters.InsertLineBreaks, saveAsParameters.AllowSubstitutions, saveAsParameters.LineEnding, saveAsParameters.AddBiDiMaks);

            //doc.SaveAs(urlFileName, fileFormat, Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value,  Missing.Value);
            doc.Close();
        }


    }
}
