using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Oracle.ManagedDataAccess.Client;

namespace SharedClassLibrary.Connections
{
    public class OracleQueryParamCollection : IQueryParamCollection
    {
        public char EscapeChar;

        public OracleQueryParamCollection()
        {
            EscapeChar = ':';
            Parameters = new List<IDbDataParameter>();
            BindByName = true;
        }

        public OracleQueryParamCollection(string query)
        {
            CommandText = query;
            EscapeChar = ':';
            Parameters = new List<IDbDataParameter>();
            BindByName = true;
        }
        public OracleQueryParamCollection(string query, List<IDbDataParameter> parameters, char escapeChar = ':')
        {
            CommandText = query;
            EscapeChar = escapeChar;
            Parameters = parameters;
            BindByName = true;
        }
        
        private static List<OracleParameter> CastToOracleParameters(List<IDbDataParameter> parameters)
        {
            var list = new List<OracleParameter>();
            foreach (var param in parameters)
            {
                list.Add((OracleParameter) param);
            }

            return list;
        }

        private static List<IDbDataParameter> CastToIDbDataParameters(List<OracleParameter> parameters)
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

            if (BindByName)
            {
                foreach (var p in Parameters)
                {
                    var paramValue = GetParameterStringValue(p);
                    result = result.Replace(EscapeChar + p.ParameterName.ToString(), paramValue);
                }
            }
            else
            {
                Regex regex = new Regex(@":\w+");
                foreach (var p in Parameters)
                {
                    var paramValue = GetParameterStringValue(p);
                    result = regex.Replace(result, paramValue, 1);
                }
            }

            return result;
        }

        private string GetParameterStringValue(IDbDataParameter p)
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

            return paramValue;
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
        public bool BindByName { get; set; }

        public List<OracleParameter> OracleParameters
        {
            get => CastToOracleParameters(Parameters);
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
