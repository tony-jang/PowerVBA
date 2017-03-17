using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;

namespace PowerVBA.Core.Connector
{
    public class SQLiteConnector
    {
        public SQLiteConnector()
        {
            SQLiteConnection.CreateFile("testdb.db");
            string strConn = @"Data Source = testdb.db";
            var dbConnection = new SQLiteConnection(strConn);
            dbConnection.Open();
        }
    }
}
