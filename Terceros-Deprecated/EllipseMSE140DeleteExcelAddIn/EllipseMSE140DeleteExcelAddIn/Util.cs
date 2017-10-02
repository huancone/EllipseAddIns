using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseMSE140DeleteExcelAddIn
{
    public class Util
    {
        public static string getEllipseDate()
        {
            return DateTime.Now.ToString("yyyyMMdd");
        }

        public static string getEllipseTime()
        {
            return DateTime.Now.ToString("HHmmss");
        }
    }
}
