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

            Rules = new Dictionary<string, Func<string, bool>>();

            Create();
        }

        public Dictionary<string, Func<string, bool>> Rules { get; internal set; }

        /// <summary>
        /// 추가할 룰을 추가합니다. 들어오는 string에 룰을 지정합니다. ruleName이 빈칸이면 False를 반환합니다.
        /// </summary>
        /// <param name="expression"></param>
        public bool AddRule(Func<string,bool> expression, string ruleName)
        {
            if (ruleName == "")
                return false;

            if (Rules.ContainsKey(ruleName))
                return false;

            Rules[ruleName] = expression;

            return true;
        }

        public bool RemoveRule(string ruleName)
        {
            if (Rules.ContainsKey(ruleName))
                Rules.Remove(ruleName);
            else
                return false;

            return true;
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
        /// <param name="checkItem">추가할 아이템이 존재하는지 체크합니다. 존재시 True를 반환합니다.</param>
        /// <returns></returns>
        public bool Add(string data, bool checkItem = true)
        {
            if (!connector.IsEnabled)
                return false;

            if (Contains(data))
                return true;

            foreach(var exp in Rules)
            {
                if (!exp.Value.Invoke(data))
                    return false;
            }

            query = $"INSERT INTO {TableName} (Data) VALUES (?)";

            SQLiteParameter param = dbCommand.CreateParameter();
            param.Value = data;

            dbCommand.CommandText = query;
            dbCommand.Parameters.Add(param);

            int InsertRow = dbCommand.ExecuteNonQuery();

            dbCommand.Parameters.Clear();

            return (InsertRow == 1);
        }

        public bool Contains(string data)
        {
            if (!connector.IsEnabled)
                return false;

            query = $"SELECT * FROM {TableName} WHERE Data = \"{data}\";";
            dbCommand.CommandText = query;

            using (SQLiteDataReader SQLiteReader = dbCommand.ExecuteReader())
            {
                while (SQLiteReader.Read())
                {
                    string str = SQLiteReader.GetValue(0).ToString();
                    return true;
                }
            }

            return false;
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