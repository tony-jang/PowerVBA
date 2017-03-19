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

            
            string strConn = @"Data Source = testdb.db";
            dbConnection = new SQLiteConnection(strConn);
            dbConnection.Open();

            if (dbConnection.State != System.Data.ConnectionState.Open) return;


            // RecentFileTable이 없을때에는 생성한다.
            strQuery = "CREATE TABLE IF NOT EXISTS RecentFileTable ( Location  VARCHAR(100) " +
                ", Type  INT )";

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

            // PreDeclareFunction  없을때에는 생성한다.
            strQuery = "CREATE TABLE IF NOT EXISTS PreDeclareFunction ( Code  VARCHAR(100000) " +
                ", Caption  VARCHAR(100) )";

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
        public bool RecentFile_Add(string FileLocation, PresentationType type)
        {
            if (!Enabled) return false;
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
        public bool RecentFile_Remove(string FileLocation)
        {
            if (!Enabled) return false;

            return true;
        }
        
        /// <summary>
        /// 최근 파일 목록을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<string> RecentFile_Get()
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
                    string type = SQLiteReader.GetValue(1).ToString();
                    strList.Add(data);
                }
            }

            return strList;
        }

        public int RecentFile_Count()
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


        #region [  PreDeclared Functions  ]

        public void Func_Add(string code, string name)
        {

        }

        public void Func_Remove(string name)
        {

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
