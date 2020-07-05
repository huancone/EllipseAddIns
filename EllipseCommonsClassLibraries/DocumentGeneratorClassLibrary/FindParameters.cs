using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace DocumentGeneratorClassLibrary
{
    public class FindParameters
    {
        public bool MatchCase = false;
        public bool MatchWholeWord = true;
        public bool MatchWildCards = false;
        public bool MatchSoundsLike = false;
        public bool MatchAllWordForms = false;
        public bool Forward = true;
        public bool Format = false;
        public bool MatchKashida = false;
        public bool MatchDiacritics = false;
        public bool MatchControl = false;
        public bool ReadOnly = false;
        public bool Visible = true;
        public WdReplace Replace = WdReplace.wdReplaceAll;
        public WdFindWrap Wrap = WdFindWrap.wdFindContinue;
        public bool MatchAlefHamza = true;
    }
}
