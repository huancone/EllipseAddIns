using System;
using SharedClassLibrary.Ellipse;

namespace EllipseSAO900AddIn
{
    /// <summary>
    ///     Crea un objeto con los datos de la transaccion contable, se contruye con la informacion del distrito y el
    ///     transaction group key, que busca esta llave en la tabla msf900
    /// </summary>
    public class Transaction
    {

        public Transaction(EllipseFunctions ef, string districtCode, string transactionNo)
        {
            try
            {
                if (string.IsNullOrEmpty(districtCode) || string.IsNullOrEmpty(transactionNo))
                {
                    Error = "Transaccion Invalida";
                    return;
                }
                var sqlQuery = Queries.GetTransactionInfo(districtCode, transactionNo, ef.DbReference,
                    ef.DbLink);

                var drTransactionNo = ef.GetQueryResult(sqlQuery);

                if (drTransactionNo != null && !drTransactionNo.IsClosed)
                {
                    while (drTransactionNo.Read())
                    {
                        FullPeriod = drTransactionNo["FULL_PERIOD"].ToString();
                        Account = drTransactionNo["ACCOUNT_CODE"].ToString();
                        ProjectNo = drTransactionNo["PROJECT_NO"].ToString();
                        Ind = drTransactionNo["IND"].ToString();
                        TranAmount = drTransactionNo["TRAN_AMOUNT"].ToString();
                        TranAmountS = drTransactionNo["TRAN_AMOUNT_S"].ToString();
                    }
                }
                else
                {
                    Error = "La Transaccion no Existe";
                }
            }
            catch (Exception error)
            {
                Error = error.Message;
            }
        }

        public string FullPeriod { get; private set; }
        public string Account { get; private set; }
        public string ProjectNo { get; private set; }
        public string Ind { get; private set; }
        public string TranAmount { get; private set; }
        public string TranAmountS { get; private set; }
        public string Error { get; private set; }
    }

}
