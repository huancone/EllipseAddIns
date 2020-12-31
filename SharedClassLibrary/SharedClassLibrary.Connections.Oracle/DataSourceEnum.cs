using System;
using System.Data;
using System.Data.Common;
using Oracle.ManagedDataAccess.Client;

namespace SharedClassLibrary.Connections.Oracle
{
    public class DataSourceEnum
    {
        public static void Main()
        {
            OracleClientFactory factory = OracleClientFactory.Instance;
            
            if (factory.CanCreateDataSourceEnumerator)
            {

                var dsenum = factory.CreateDataSourceEnumerator() as OracleDataSourceEnumerator;
                
                DataTable dt = dsenum.GetDataSources();
                foreach (DataRow row in dt.Rows)
                {
                    Console.WriteLine(dt.Columns[0] + " : " + row[0]);
                    Console.WriteLine(dt.Columns[1] + " : " + row[1]);
                    Console.WriteLine(dt.Columns[2] + " : " + row[2]);
                    Console.WriteLine(dt.Columns[3] + " : " + row[3]);
                    Console.WriteLine(dt.Columns[4] + " : " + row[4]);
                    Console.WriteLine("--------------------");


                }
            }
        }
    }
}
