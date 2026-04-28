using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ChuyenlieulotbbNEW
{
    public class SqlCnn
    {
        private const int DefaultCommandTimeoutSeconds = 6000;

        public static DataTable ExecuteQuery(string query, string ConnectionString, CommandType commandType = CommandType.Text,
           Dictionary<string, object> param = null)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = CreateCommand(query, conn, commandType, param))
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        return dt;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    return new DataTable();
                }
                finally
                {
                    if (conn.State != ConnectionState.Closed)
                        conn.Close();
                }
            }
        }

        public static bool ExecuteNonQuery(string query, string ConnectionString, CommandType commandType = CommandType.Text, Dictionary<string, object> param = null)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = CreateCommand(query, conn, commandType, param))
                    {
                        int affectedRows = cmd.ExecuteNonQuery();
                        return affectedRows > 0;
                    }
                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine("SQL Error: " + sqlEx.Message);
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    return false;
                }
                finally
                {
                    if (conn.State != ConnectionState.Closed)
                    {
                        conn.Close();
                    }
                }
            }
        }

        private static SqlCommand CreateCommand(
            string query,
            SqlConnection conn,
            CommandType commandType,
            Dictionary<string, object> param)
        {
            SqlCommand cmd = new SqlCommand(query, conn)
            {
                CommandType = commandType,
                CommandTimeout = DefaultCommandTimeoutSeconds
            };

            AddParameters(cmd, param);
            return cmd;
        }

        private static void AddParameters(SqlCommand cmd, Dictionary<string, object> param)
        {
            if (param == null || param.Count == 0)
            {
                return;
            }

            foreach (var item in param)
            {
                var parameter = new SqlParameter("@" + item.Key, SqlDbType.NVarChar, 50)
                {
                    Value = item.Value ?? DBNull.Value
                };
                cmd.Parameters.Add(parameter);
            }
        }

    }
}
