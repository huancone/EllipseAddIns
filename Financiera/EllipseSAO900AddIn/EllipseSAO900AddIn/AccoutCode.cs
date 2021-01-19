using System;
using SharedClassLibrary.Ellipse;

namespace EllipseSAO900AddIn
{
    /// <summary>
    ///     Crea un objeto con la informacion de la combinacion centro/detalle verificando el estado y las variables de control
    ///     de proyecto, orden y subledger.
    /// </summary>
    public class AccountCode
    {
        public AccountCode(EllipseFunctions ef, string districtCode, string accountCode)
        {
            try
            {
                if (string.IsNullOrEmpty(districtCode) || string.IsNullOrEmpty(accountCode))
                {
                    Error = "AccoundeCode Invalida";
                    return;
                }

                accountCode = accountCode.Contains(";")
                    ? accountCode.Substring(0, accountCode.IndexOf(";", StringComparison.Ordinal))
                    : accountCode;

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
                        Error = (drAccountCode["ACTIVE_STATUS"].ToString() == "I") ? "AccountCode Inactivo" : null;
                    }



                }
                else
                {
                    Error = "Centro de Costos Destino No Valido";
                }
            }
            catch (Exception error)
            {
                Error = error.Message;
            }
        }

        public string Error { get; private set; }
        public string ActiveStatus { get; private set; }
        public string ProjectEntriInd { get; private set; }
        public string WorkOrderEntryInd { get; private set; }
        public string SubLedgerInd { get; private set; }
    }
}
