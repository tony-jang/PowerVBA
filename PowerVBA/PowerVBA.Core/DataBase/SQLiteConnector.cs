using System.Collections.Generic;
using System.Linq;
using System.Data.SQLite;
using System.Diagnostics;
using System.Windows;
using System.IO;

namespace PowerVBA.Core.DataBase
{
    public class SQLiteConnector
    {
        internal SQLiteConnection dbConnection;
        internal bool IsEnabled { get; set; }


        public SQLiteTable FileTable { get; set; }
        public SQLiteTable FolderTable { get; set; }


        public SQLiteConnector()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            
            string strConn = @"Data Source = PowerVBA.db";
            dbConnection = new SQLiteConnection(strConn);
            dbConnection.Open();
            
            IsEnabled = (dbConnection.State == System.Data.ConnectionState.Open);

            FileTable = new SQLiteTable(this, "FileTable", 100);
            FolderTable = new SQLiteTable(this, "FolderTable", 100);
        }
        
        
    }
}
