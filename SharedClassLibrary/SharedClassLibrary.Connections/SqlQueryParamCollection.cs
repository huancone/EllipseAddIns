using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace SharedClassLibrary.Connections
{
    public class SqlQueryParamCollection : IQueryParamCollection
    {
        public char EscapeChar;

        public SqlQueryParamCollection()
        {
            EscapeChar = '@';
            Parameters = new List<IDbDataParameter>();
        }

        public SqlQueryParamCollection(string query)
        {
            CommandText = query;
            EscapeChar = '@';
            Parameters = new List<IDbDataParameter>();
        }
        public SqlQueryParamCollection(string query, List<IDbDataParameter> parameters, char escapeChar = '@')
        {
            CommandText = query;
            EscapeChar = escapeChar;
            Parameters = parameters;
        }
        
        private static List<SqlParameter> CastToSqlParameters(List<IDbDataParameter> parameters)
        {
            var list = new List<SqlParameter>();
            foreach (var param in parameters)
            {
                list.Add((SqlParameter) param);
            }

            return list;
        }

        private static List<IDbDataParameter> CastToIDbDataParameters(List<SqlParameter> parameters)
        {
            var list = new List<IDbDataParameter>();
            foreach (var param in parameters)
            {
                list.Add(param);
            }

            return list;
        }
        public string GetGeneratedSql()
        {
            var result = CommandText;
            if (Parameters == null)
                return result;
            foreach (var p in Parameters)
            {
                string paramValue;
                switch (p.Value)
                {
                    case string _:
                        paramValue = "'" + p.Value?.ToString() + "'";
                        break;
                    case DateTime _:
                        paramValue = Convert.ToDateTime(p.Value).ToString("yyyyMMdd hhmmss tt");
                        paramValue = "TO_DATE('" + paramValue + "', 'YYYYMMDD HHMISS AM')";
                        break;
                    default:
                        paramValue = p.Value?.ToString();
                        break;
                }

                result = result.Replace(EscapeChar + p.ParameterName.ToString(), paramValue);
            }
            return result;
        }

        public string UpdatedEscapedCharacter(char newChar)
        {
            var result = CommandText;
            if (Parameters == null)
                return result;

            var oldChar = EscapeChar;
            foreach (var p in Parameters)
                result = result.Replace(oldChar + p.ParameterName.ToString(), newChar + p.ParameterName.ToString());

            CommandText = result;
            return result;
        }

        public List<IDbDataParameter> Parameters { get; set; }

        public string CommandText { get; set; }
        public bool BindByName { get; set; }//N/A

        public List<SqlParameter> SqlParameters
        {
            get => CastToSqlParameters(Parameters);
            set => Parameters = CastToIDbDataParameters(value);
        }
        public void AddParam(IDbDataParameter parameter)
        {
            if (Parameters == null)
                Parameters = new List<IDbDataParameter>();
            Parameters.Add(parameter);
        }

    }
}
