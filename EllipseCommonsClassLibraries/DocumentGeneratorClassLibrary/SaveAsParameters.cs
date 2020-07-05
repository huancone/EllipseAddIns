using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace DocumentGeneratorClassLibrary
{
    public class SaveAsParameters
    {
        public object LockComments = Missing.Value; //null or boolean
        public object Password = Missing.Value; //null or string
        public object AddToRecentFilesfalse = Missing.Value; //null or boolean
        public object WritePassword = Missing.Value; //null or string
        public object ReadOnlyRecommended = Missing.Value; //null or boolean
        public object EmbedTrueTypeFonts = Missing.Value; //null or boolean
        public object SaveNativePictureFormat = Missing.Value; //null or boolean
        public object SaveFormsData = Missing.Value; //null or boolean
        public object SaveAsAoceLetter = Missing.Value; //null or boolean
        public object Encoding = Missing.Value; //null or MSOEncoding
        public object InsertLineBreaks = Missing.Value;
        public object AllowSubstitutions = Missing.Value;
        public object LineEnding = Missing.Value; //null or WdLineEndingType
        public object AddBiDiMaks = Missing.Value; //null or boolean
    }
}
