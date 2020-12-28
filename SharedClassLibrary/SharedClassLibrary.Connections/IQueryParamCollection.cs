using System.Collections.Generic;
using System.Data;


namespace SharedClassLibrary.Connections
{
    public interface IQueryParamCollection
    {
        List<IDbDataParameter> Parameters { get; set; }
        string CommandText { get; set; }
        bool BindByName { get; set; }
        void AddParam(IDbDataParameter parameter);
        string GetGeneratedSql();
        string UpdatedEscapedCharacter(char newChar);
    }
}