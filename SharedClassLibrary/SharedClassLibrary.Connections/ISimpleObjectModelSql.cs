using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SharedClassLibrary.Connections
{
    public interface ISimpleObjectModelSql
    {
        void SetFromDataRecord(IDataRecord dr);
    }
}
