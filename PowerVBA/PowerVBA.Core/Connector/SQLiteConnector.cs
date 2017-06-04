using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Diagnostics;
using System.Windows;
using System.IO;

namespace PowerVBA.Core.Connector
{
    public class SQLiteConnector
    {
        internal string strQuery = "";
        internal SQLiteConnection dbConnection;
        internal SQLiteCommand dbCommand;
        internal bool Enabled = false;
        public SQLiteConnector()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            
            string strConn = @"Data Source = PowerVBA.db";
            dbConnection = new SQLiteConnection(strConn);
            dbConnection.Open();

            if (dbConnection.State != System.Data.ConnectionState.Open) return;


            // RecentFileTable이 없을때에는 생성한다.
            strQuery = "CREATE TABLE IF NOT EXISTS RecentFileTable ( Location  VARCHAR(100) )";

            dbCommand = dbConnection.CreateCommand();
            dbCommand.CommandText = strQuery;

            try
            {
                dbCommand.ExecuteNonQuery();
                Enabled = true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.ToString());
                Enabled = false;
                throw;
            }
            
            // DllLocation 테이블이 없을때에는 생성한다.
            strQuery = "CREATE TABLE IF NOT EXISTS DllLocation ( Location VARCHAR(255) " +
                ", Name VARCHAR(100))";

            dbCommand = dbConnection.CreateCommand();
            dbCommand.CommandText = strQuery;

            try
            {
                dbCommand.ExecuteNonQuery();
                Enabled = true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.ToString());
                Enabled = false;
                throw;
            }

            // RecentFolder 테이블이 없을때에는 생성한다.
            strQuery = "CREATE TABLE IF NOT EXISTS RecentFolder ( Location VARCHAR(100) )";

            dbCommand = dbConnection.CreateCommand();
            dbCommand.CommandText = strQuery;

            try
            {
                dbCommand.ExecuteNonQuery();
                Enabled = true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.ToString());
                Enabled = false;
                throw;
            }

        }
        
        #region [  Recent Files  ]

        /// <summary>
        /// 최근 파일 목록에 추가합니다.
        /// </summary>
        /// <param name="FileLocation">추가할 파일 위치입니다.</param>
        /// <returns></returns>
        public bool RecentFileAdd(string FileLocation)
        {
            if (!Enabled) return false;
            if (RecentFileContains(FileLocation)) return true;
            strQuery = "INSERT INTO RecentFileTable (Location) VALUES (?)";
            dbCommand.CommandText = strQuery;
            dbCommand.Parameters.Clear();

            SQLiteParameter paramLoc = dbCommand.CreateParameter();

            paramLoc.Value = FileLocation;

            dbCommand.Parameters.Add(paramLoc);

            // 1 : 성공       이외 : 실패
            int InsertRow = dbCommand.ExecuteNonQuery();

            return (InsertRow == 1);
        }

        /// <summary>
        /// 최근 파일 목록에서 제거합니다.
        /// </summary>
        /// <param name="FileLocation">제거할 파일 위치입니다.</param>
        /// <returns></returns>
        public bool RecentFileRemove(string FileLocation)
        {
            if (!Enabled) return false;

            strQuery = $"DELETE FROM RecentFileTable WHERE Location = \"{FileLocation}\";";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();

            int DelRow = dbCommand.ExecuteNonQuery();

            return (DelRow == 1);
        }

        public bool RecentFileContains(string FileLocation)
        {
            var strList = new List<string>();
            if (!Enabled) return false;
            strQuery = $"SELECT * FROM RecentFileTable WHERE Location = \"{FileLocation}\";";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                {
                    string data = SQLiteReader.GetValue(0).ToString();
                    strList.Add(data);
                }
            }

            return strList.Count != 0;
        }
        
        /// <summary>
        /// 최근 파일 목록을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<string> RecentFileGet()
        {
            var strList = new List<string>();
            if (!Enabled) return strList;
            strQuery = "SELECT * FROM RecentFileTable;";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();
            
            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                {
                    string data = SQLiteReader.GetValue(0).ToString();
                    strList.Add(data);
                }
            }

            return strList;
        }

        public int RecentFileCount()
        {
            if (!Enabled) return 0;

            strQuery = "SELECT COUNT (*) FROM RecentFileTable;";

            dbCommand.Parameters.Clear();
            dbCommand.CommandText = strQuery;

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                if (SQLiteReader.Read())
                {
                    return int.Parse(SQLiteReader.GetValue(0).ToString());
                }
                return 0;
            }
        }

        #endregion
        
        #region [  Recent Folder  ]

        /// <summary>
        /// 최근 폴더 목록에 추가합니다.
        /// </summary>
        /// <param name="FolderLocation">추가할 파일 위치입니다.</param>
        /// <returns></returns>
        public bool RecentFolderAdd(string FolderLocation)
        {
            if (!Enabled) return false;
            if (RecentFolderContains(FolderLocation)) return true;
            strQuery = "INSERT INTO RecentFolder (Location) VALUES (?)";
            dbCommand.CommandText = strQuery;
            dbCommand.Parameters.Clear();

            SQLiteParameter paramLoc = dbCommand.CreateParameter();

            paramLoc.Value = FolderLocation;

            dbCommand.Parameters.Add(paramLoc);

            // 1 : 성공       이외 : 실패
            int InsertRow = dbCommand.ExecuteNonQuery();

            return (InsertRow == 1);
        }

        /// <summary>
        /// 최근 폴더 목록에서 제거합니다.
        /// </summary>
        /// <param name="FolderLocation">제거할 파일 위치입니다.</param>
        /// <returns></returns>
        public bool RecentFolderRemove(string FolderLocation)
        {
            if (!Enabled) return false;

            strQuery = $"DELETE FROM RecentFolder WHERE Location = \"{FolderLocation}\";";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();

            int DelRow = dbCommand.ExecuteNonQuery();

            return (DelRow == 1);
        }

        public bool RecentFolderContains(string FolderLocation)
        {
            var strList = new List<string>();
            if (!Enabled) return false;
            strQuery = $"SELECT * FROM RecentFolder WHERE Location = \"{FolderLocation}\";";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                {
                    string data = SQLiteReader.GetValue(0).ToString();
                    strList.Add(data);
                }
            }

            return strList.Count != 0;
        }

        /// <summary>
        /// 최근 폴더 목록을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<string> RecentFolderGet()
        {
            var strList = new List<string>();
            if (!Enabled) return strList;
            strQuery = "SELECT * FROM RecentFolder;";

            dbCommand.CommandText = strQuery;

            dbCommand.Parameters.Clear();

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                {
                    string data = SQLiteReader.GetValue(0).ToString();
                    strList.Add(data);
                }
            }

            return strList;
        }

        public int RecentFolderCount()
        {
            if (!Enabled) return 0;

            strQuery = "SELECT COUNT (*) FROM RecentFolder;";

            dbCommand.Parameters.Clear();
            dbCommand.CommandText = strQuery;

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                if (SQLiteReader.Read())
                {
                    return int.Parse(SQLiteReader.GetValue(0).ToString());
                }
                return 0;
            }
        }

        #endregion


        public enum PresentationType
        {
            /// <summary>
            /// 기본 프레젠테이션
            /// </summary>
            General,
            /// <summary>
            /// VBA 파일을 사용하는 프레젠테이션
            /// </summary>
            VBAAssist,
            /// <summary>
            /// 가상 프레젠테이션
            /// </summary>
            Virtual
        }

    }
}
