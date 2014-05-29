using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Data.Common;
using System.Globalization;

namespace SIStation
{
    #region 委托

    /// <summary>
    /// 数据库操作命令委托
    /// </summary>
    /// <param name="dbCommand">操作命令</param>
    /// <returns>委托的命令</returns>
    internal delegate object CommandDelegate(SQLiteCommand dbCommand);

    /// <summary>
    /// DbDataReader命令委托
    /// </summary>
    /// <param name="dbDataReader">DbDataReader</param>
    internal delegate void DbDataReadDelegate(SQLiteDataReader dbDataReader);

    #endregion

    class SQLiteHelper
    {
        #region 字段

        /// <summary>
        /// 数据库操作辅助类实例对象
        /// </summary>
        private static SQLiteHelper _Instance;

        /// <summary>
        /// 表示连接字符串配置文件节中的单个命名连接字符串
        /// </summary>
        private ConnectionStringSettings _ConnectionStringSettings;

        /// <summary>
        /// 保存数据库链接字符串
        /// </summary>
        public static string DataConnectionString = "ConnectionString";

        #endregion


                #region 构造函数

        /// <summary>
        /// 无参构造函数
        /// </summary>
        private SQLiteHelper()
        {
            _ConnectionStringSettings = ConfigurationManager.ConnectionStrings[DataConnectionString];
        }

        #endregion


        #region 类属性

        /// <summary>
        /// 单实例
        /// </summary>
        public static SQLiteHelper Instance
        {
            get
            {
                _Instance = _Instance ?? new SQLiteHelper();
                return _Instance;
            }
        }

        #endregion


        #region 方法

        #region CreateConnection

        /// <summary>
        /// 创建一个数据库连接
        /// </summary>
        /// <returns>数据库连接对象</returns>
        public SQLiteConnection CreateConnection()
        {
            string dbPath = Path.Combine(Directory.GetCurrentDirectory(), _ConnectionStringSettings.ConnectionString);
            if (!File.Exists(dbPath))
            {
                string dir = Path.GetDirectoryName(dbPath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                using (File.Create(dbPath)) { };
            }
            string connectionString = "Data Source=" + dbPath;
            SQLiteConnection sqliteConnection = new SQLiteConnection(connectionString);
            return sqliteConnection;
        }

        #endregion

        #region ExecuteReader

        /// <summary>
        /// 读取Sql的内容倒DataTable
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">查询参数</param>
        /// <returns>DataTable</returns>
        public DataTable ExecuteReader(string sql, IList<DbParameter> para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    DataTable dt = new DataTable();
                    dt.Locale = CultureInfo.InvariantCulture;
                    dt.Load(dr);
                    return dt;
                }
            };
            return (DataTable)ExecuteCmdCallback(CommandType.Text, sql, cd, para);
        }
        /// <summary>
        /// 根据Sql创建一个DataTable的结果集
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable ExecuteReader(string sql, params DbParameter[] para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    DataTable dt = new DataTable();
                    dt.Locale = CultureInfo.InvariantCulture;
                    dt.Load(dr);
                    return dt;
                }
            };
            return (DataTable)ExecuteCmdCallback(CommandType.Text, sql, cd, para);
        }

        #endregion

        #region ExecuteNonQuery

        /// <summary>
        /// 执行无结果集Sql
        /// </summary>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">占位符参数</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteNonQuery(string sql, params DbParameter[] para)
        {
            return ExecuteNonQuery(CommandType.Text, sql, para);
        }

        /// <summary>
        /// 执行无结果集Sql
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">占位符参数</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteNonQuery(CommandType commandType, string sql, params DbParameter[] para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                return cmd.ExecuteNonQuery();
            };
            return (int)ExecuteCmdCallback(commandType, sql, cd, para);
        }

        /// <summary>
        /// 执行无结果集Sql
        /// </summary>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">占位符参数</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteNonQuery(string sql, IList<DbParameter> para)
        {
            return ExecuteNonQuery(CommandType.Text, sql, para);
        }

        /// <summary>
        /// 执行无结果集Sql
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">占位符参数</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteNonQuery(CommandType commandType, string sql, IList<DbParameter> para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                return cmd.ExecuteNonQuery();
            };
            return (int)ExecuteCmdCallback(commandType, sql, cd, para);
        }

        #endregion

        #region CreateDataSet

        /// <summary>
        /// 根据Sql语句创建一个DataSet类型的结果集
        /// </summary>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">查询参数</param>
        /// <returns>DataSet结果集</returns>
        public DataSet CreateDataSet(string sql, params DbParameter[] para)
        {
            return CreateDataSet(CommandType.Text, sql, para);
        }

        /// <summary>
        /// 根据Sql语句创建一个DataSet类型的结果集
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">查询参数</param>
        /// <returns>DataSet结果集</returns>
        public DataSet CreateDataSet(CommandType commandType, string sql, params DbParameter[] para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                using(SQLiteDataAdapter da = new SQLiteDataAdapter(sql, cmd.Connection))
                {
                    DataSet ds = new DataSet();
                    ds.Locale = CultureInfo.InvariantCulture;
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    return ds;
                }
            };
            return (DataSet)ExecuteCmdCallback(commandType, sql, cd, para);
        }

        /// <summary>
        /// 根据Sql语句创建一个DataSet类型的结果集
        /// </summary>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">查询参数</param>
        /// <returns>DataSet结果集</returns>
        public DataSet CreateDataSet(string sql, IList<DbParameter> para)
        {
            return CreateDataSet(CommandType.Text, sql, para);
        }

        /// <summary>
        /// 根据Sql语句创建一个DataSet类型的结果集
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="para">查询参数</param>
        /// <returns>DataSet结果集</returns>
        public DataSet CreateDataSet(CommandType commandType, string sql, IList<DbParameter> para)
        {
            CommandDelegate cd = delegate(SQLiteCommand cmd)
            {
                using (SQLiteDataAdapter da = new SQLiteDataAdapter(sql, cmd.Connection))
                {
                    DataSet ds = new DataSet();
                    ds.Locale = CultureInfo.InvariantCulture;
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    return ds;
                }
            };
            return (DataSet)ExecuteCmdCallback(commandType, sql, cd, para);
        }

        #endregion

        #region ExecuteCmdCallback

        /// <summary>
        /// 执行带参数与委托命令的查询语句，并返回相关的委托命令
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="commandDelegate">委托类型</param>
        /// <param name="para">参数集合</param>
        /// <returns>委托的命令</returns>
        private object ExecuteCmdCallback(CommandType commandType, string sql, CommandDelegate commandDelegate, params DbParameter[] para)
        {
            using (SQLiteConnection sqliteCon = CreateConnection())
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    cmd.CommandType = commandType;
                    cmd.CommandText = sql;
                    cmd.Connection = sqliteCon;

                    foreach (DbParameter dp in para)
                    {
                        cmd.Parameters.Add(dp);
                    }
                    sqliteCon.Open();
                    return commandDelegate(cmd);
                }
            }
        }

        /// <summary>
        /// 执行带参数与委托命令的查询语句，并返回相关的委托命令
        /// </summary>
        /// <param name="commandType">命令类型</param>
        /// <param name="sql">Sql语句</param>
        /// <param name="commandDelegate">委托类型</param>
        /// <param name="para">参数集合</param>
        /// <returns>委托的命令</returns>
        private object ExecuteCmdCallback(CommandType commandType, string sql, CommandDelegate commandDelegate, IList<DbParameter> para)
        {
            using (SQLiteConnection sqliteCon = CreateConnection())
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    cmd.CommandType = commandType;
                    cmd.CommandText = sql;
                    cmd.Connection = sqliteCon;

                    foreach (DbParameter dp in para)
                    {
                        cmd.Parameters.Add(dp);
                    }
                    sqliteCon.Open();
                    return commandDelegate(cmd);
                }
            }
        }

        #endregion

        #endregion

    }
}
