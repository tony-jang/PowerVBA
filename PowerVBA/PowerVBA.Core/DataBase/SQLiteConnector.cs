using System.Data.SQLite;
using System.Diagnostics;
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
            
            string strConn = @"Data Source = PowerVBA.db;Password=POWERVBADATABASE";
            dbConnection = new SQLiteConnection(strConn);
            dbConnection.Open();
            
            IsEnabled = (dbConnection.State == System.Data.ConnectionState.Open);

            FileTable = new SQLiteTable(this, "FileTable", 100);
            FileTable.AddRule(str => new FileInfo(str).Exists, "ExistsCheck");

            FolderTable = new SQLiteTable(this, "FolderTable", 100);
            FolderTable.AddRule(str => new DirectoryInfo(str).Exists, "ExistsCheck");
        }
        
        
    }
}
