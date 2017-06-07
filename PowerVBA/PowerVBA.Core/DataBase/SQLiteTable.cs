using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Core.DataBase
{
    /// <summary>
    /// 문자 (VARCHAR/String) 하나를 필드로 사용하는 테이블을 관리합니다.
    /// </summary>
    public class SQLiteTable 
    {
        public SQLiteTable(SQLiteConnector connector, string tableName, int size = 150)
        {
            this.connector = connector;
            this.size = size;
            this.TableName = tableName;

            Create();
        }

        private SQLiteConnector connector;
        private int size;
        private SQLiteCommand dbCommand;
        private string query;

        public string TableName { get; set; }
        

        /// <summary>
        /// 아이템들을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public List<string> Get()
        {
            var strList = new List<string>();

            if (!connector.IsEnabled)
                return strList;

            query = $"SELECT * FROM {TableName};";
            dbCommand.CommandText = query;

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                    strList.Add(SQLiteReader.GetValue(0).ToString());
            }

            return strList;
        }

        /// <summary>
        /// 아이템의 갯수를 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public int Count()
        {
            if (!connector.IsEnabled)
                return 0;

            query = $"SELECT COUNT (*) FROM {TableName};";
            dbCommand.CommandText = query;

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                if (SQLiteReader.Read())
                    return int.Parse(SQLiteReader.GetValue(0).ToString());

                return 0;
            }
        }

        /// <summary>
        /// 아이템을 추가합니다.
        /// </summary>
        /// <param name="data">추가할 아이템입니다.</param>
        /// <returns></returns>
        public bool Add(string data)
        {
            if (!connector.IsEnabled)
                return false;

            query = $"INSERT INTO {TableName} (Data) VALUES (?)";

            SQLiteParameter param = dbCommand.CreateParameter();
            param.Value = data;

            dbCommand.CommandText = query;
            dbCommand.Parameters.Add(param);

            int InsertRow = dbCommand.ExecuteNonQuery();

            dbCommand.Parameters.Clear();

            return (InsertRow == 1);
        }

        /// <summary>
        /// 아이템을 삭제합니다.
        /// </summary>
        /// <param name="data">삭제할 아이템입니다.</param>
        /// <returns></returns>
        public bool Remove(string data)
        {
            if (!connector.IsEnabled)
                return false;

            query = $"DELETE FROM {TableName} WHERE Data = \"{data}\" COLLATE NOCASE";

            dbCommand.CommandText = query;
            int DelRow = dbCommand.ExecuteNonQuery();

            return (DelRow > 0);
        }

        /// <summary>
        /// 테이블을 초기화합니다.
        /// </summary>
        /// <returns></returns>
        public bool Clear()
        {
            if (!connector.IsEnabled)
                return false;

            query = $"DELETE FROM {TableName};";
            
            dbCommand.CommandText = query;

            int DelRow = dbCommand.ExecuteNonQuery();

            return (DelRow != 0);
        }

        /// <summary>
        /// 테이블을 생성합니다.
        /// </summary>
        /// <returns></returns>
        public bool Create()
        {
            var query = $"CREATE TABLE IF NOT EXISTS {TableName} ( Data VARCHAR({size}) )";

            dbCommand = connector.dbConnection.CreateCommand();
            dbCommand.CommandText = query;

            try
            {
                dbCommand.ExecuteNonQuery();
                return true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        }
    }
}