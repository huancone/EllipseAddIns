using System;
using SharedClassLibrary.Ellipse;

namespace EllipseMSO905ExcelAddIn
{
    public class AccountCode
    {
        public AccountCode(EllipseFunctions ef, string districtCode, string accountCode)
        {
            try
            {
                if (string.IsNullOrEmpty(districtCode) || string.IsNullOrEmpty(accountCode))
                {
                    Error = "Account Code Invalida";
                    return;
                }

                if (accountCode.Contains(";"))
                {
                    Mnemonic = accountCode.Contains(";")
                    ? accountCode.Substring(accountCode.IndexOf(";", StringComparison.Ordinal) + 1, accountCode.Length - accountCode.IndexOf(";", StringComparison.Ordinal) - 1).Replace("=", "")
                    : "";
                    accountCode = accountCode.Contains(";")
                        ? accountCode.Substring(0, accountCode.IndexOf(";", StringComparison.Ordinal))
                        : accountCode;



                    var mnemonicQuery = Queries.GetSupplierMnemonic(Mnemonic, ef.DbReference, ef.DbLink);
                    var drMnemonic = ef.GetQueryResult(mnemonicQuery);
                    if (drMnemonic != null && !drMnemonic.IsClosed)
                    {
                        drMnemonic.Read();
                        if (Convert.ToInt32(drMnemonic["CANTIDAD"].ToString()) == 1)
                            Account = accountCode + ";" + drMnemonic["GL_COLLOQ_CD"];
                        else
                            Error = " No se puede determinar la cuenta";
                    }
                    else
                        Error = " Mnenomico No Valido";
                    if (drMnemonic != null) drMnemonic.Close();
                }

                var sqlQuery = Queries.GetAccountCodeInfo(districtCode, accountCode, ef.DbReference,
                    ef.DbLink);

                var drAccountCode = ef.GetQueryResult(sqlQuery);

                if (drAccountCode != null && !drAccountCode.IsClosed)
                {
                    while (drAccountCode.Read())
                    {
                        ActiveStatus = drAccountCode["ACTIVE_STATUS"].ToString();
                        ProjectEntriInd = drAccountCode["PROJ_ENTRY_IND"].ToString();
                        WorkOrderEntryInd = drAccountCode["WO_ENTRY_IND"].ToString();
                        SubLedgerInd = drAccountCode["SUBLEDGER_IND"].ToString();
                        Error = (drAccountCode["ACTIVE_STATUS"].ToString() == "I") ? ", AccountCode Inactivo" : Error;
                    }
                }
                else
                {
                    Error = " Centro de Costos No Valido";
                }
                if (drAccountCode != null) drAccountCode.Close();
            }
            catch (Exception error)
            {
                Error = error.Message;
            }
        }

        public string Error { get; set; }
        public string ActiveStatus { get; set; }
        public string Account { get; set; }
        public string ProjectEntriInd { get; set; }
        public string WorkOrderEntryInd { get; set; }
        public string SubLedgerInd { get; set; }
        public string Mnemonic { get; set; }
    }
}
