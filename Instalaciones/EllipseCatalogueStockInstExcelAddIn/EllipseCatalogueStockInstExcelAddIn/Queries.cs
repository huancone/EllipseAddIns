namespace EllipseCatalogueStockInstExcelAddIn
{
    public static class Queries
    {
        public static string GetContractData(string contractNo, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
                           "SELECT DISTINCT" +
                           "  CON.CONTRACT_NO, " +
                           "  CON.PORTION_NO, " +
                           "  CON.ELEMENT_NO, " +
                           "  CON.CATEGORY_NO, " +
                           "  PN.STOCK_CODE, " +
                           "  CON.CATEG_DESC, " +
                           "  CAT.HOME_WHOUSE, " +
                           "  WH.WHOUSE_ID " +
                           "FROM " +
                           "  ELLIPSE.MSF387 CON " +
                           "INNER JOIN " + dbReference + ".MSF110" + dbLink + " PN " +
                           "ON " +
                           "  PN.PART_NO LIKE CON.PORTION_NO || CON.ELEMENT_NO || CON.CATEGORY_NO || '%' || CON.CONTRACT_NO || '%' " +
                           "LEFT JOIN " + dbReference + ".MSF170" + dbLink + " CAT " +
                           "ON " +
                           "  PN.STOCK_CODE = CAT.STOCK_CODE " +
                           "AND CAT.DSTRCT_CODE = 'INST' " +
                           "LEFT JOIN " + dbReference + ".MSF180" + dbLink + " WH " +
                           "ON " +
                           "    CAT.STOCK_CODE = WH.STOCK_CODE " +
                           "AND CAT.DSTRCT_CODE = WH.DSTRCT_CODE " +
                           "WHERE " +
                           "  CON.CONTRACT_NO = '" + contractNo + "' " +
                           "ORDER BY " +
                           "  CON.CONTRACT_NO, " +
                           "  CON.PORTION_NO, " +
                           "  CON.ELEMENT_NO, " +
                           "  CON.CATEGORY_NO ";
            return sqlQuery;
        }
    }
}
