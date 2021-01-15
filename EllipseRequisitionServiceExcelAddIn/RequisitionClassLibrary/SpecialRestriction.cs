using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary
{
    public static class SpecialRestriction
    {
        public static List<SpecialRestrictionItem> GetPositionRestrictions(EllipseFunctions eFunctions)
        {
            var listItems = new List<SpecialRestrictionItem>();
            var query = GetSpecialRestrictionsQuery();
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drItemCodes = eFunctions.GetQueryResult(query);

            if (drItemCodes == null || drItemCodes.IsClosed) return listItems;
            while (drItemCodes.Read())
            {
                var item = new SpecialRestrictionItem
                {
                    Position = drItemCodes["POSITION"].ToString().Trim(),
                    Code = drItemCodes["TABLE_CODE"].ToString().Trim(),
                    MandatoryWorkOrder = MyUtilities.IsTrue(drItemCodes["WO_MANDATORY_FLAG"].ToString().Trim())
                };
                listItems.Add(item);
            }

            return listItems;
        }
        public class SpecialRestrictionItem
        {
            public string Position;
            public bool MandatoryWorkOrder;
            public string Code;
        }
        public static string GetSpecialRestrictionsQuery()
        {
            var query = "WITH PPP_TABLE AS" +
                        " (" +
                        "     SELECT" +
                        " SUBSTR(ASSOC_REC, 1, 1) WO_MANDATORY_FLAG," +
                        " SUBSTR(ASSOC_REC, 11, 10) POSITION_1," +
                        " SUBSTR(ASSOC_REC, 21, 10) POSITION_2," +
                        " SUBSTR(ASSOC_REC, 31, 10) POSITION_3," +
                        " SUBSTR(ASSOC_REC, 41, 10) POSITION_4," +
                        " TABLE_CODE" +
                        "     FROM ELLIPSE.MSF010" +
                        " WHERE TABLE_TYPE = '+PPP' AND ACTIVE_FLAG = 'Y' " +
                        "     )," +
                        " MODEL_FIL AS" +
                        " (" +
                        "     SELECT * FROM PPP_TABLE UNPIVOT(POSITION FOR COLUMNAS IN (POSITION_1, POSITION_2, POSITION_3, POSITION_4))" +
                        "     )" +
                        " SELECT WO_MANDATORY_FLAG, POSITION, TABLE_CODE" +
                        " FROM MODEL_FIL" +
                        " WHERE TRIM(POSITION) IS NOT NULL";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            return query;
        }
    }

}
