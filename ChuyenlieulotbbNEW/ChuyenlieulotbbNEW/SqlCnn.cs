using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChuyenlieulotbbNEW
{
    public class SqlCnn
    {
        public static DataTable ExecuteQuery(string query, string ConnectionString, CommandType commandType = CommandType.Text,
           Dictionary<string, object> param = null)
        {
            //string ConnectionString = @"Data Source=198.1.9.186;Initial Catalog=" + database + ";User ID=kendakv2;Password=kenda123";
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.CommandType = commandType;
                    if (param != null && param.Count > 0)
                        foreach (var item in param)
                        {
                            //cmd.Parameters.AddWithValue(item.Key, item.Value);
                            cmd.Parameters.Add("@" + item.Key, SqlDbType.NVarChar, 50);
                            cmd.Parameters["@" + item.Key].Value = item.Value;
                        }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    return dt;
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

        //public static bool ExecuteNonQuery(string query, string ConnectionString, CommandType commandType = CommandType.Text,
        //                            Dictionary<string, object> param = null)
        //{
        //    //string ConnectionString = @"Data Source=198.1.9.186;Initial Catalog=" + database + ";User ID=kendakv2;Password=kenda123";
        //    using (var conn = new SqlConnection(ConnectionString))
        //    {
        //        try
        //        {
        //            conn.Open();
        //            SqlCommand cmd = new SqlCommand(query, conn)
        //            {
        //                CommandType = commandType,
        //                CommandTimeout = 120 // Set the command timeout to 120 seconds (or adjust as needed)
        //            };

        //            if (param != null && param.Count > 0)
        //                foreach (var item in param)
        //                {
        //                    //cmd.Parameters.AddWithValue(item.Key, item.Value);
        //                    cmd.Parameters.Add("@" + item.Key, SqlDbType.NVarChar, 50);
        //                    cmd.Parameters["@" + item.Key].Value = item.Value;
        //                }


        //            int effectedRow = cmd.ExecuteNonQuery();
        //            return effectedRow > 0;
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.ToString());
        //            return false;
        //        }
        //        finally
        //        {
        //            if (conn.State != ConnectionState.Closed)
        //                conn.Close();
        //        }
        //    }
        //}
        public static bool ExecuteNonQuery(string query, string ConnectionString, CommandType commandType = CommandType.Text, Dictionary<string, object> param = null)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = commandType;
                        cmd.CommandTimeout = 6000; // Tăng lên nếu cần

                        if (param != null)
                        {
                            foreach (var item in param)
                            {
                                // Chỉ rõ kiểu dữ liệu và kích thước của tham số
                                var parameter = new SqlParameter("@" + item.Key, SqlDbType.NVarChar, 50)
                                {
                                    Value = item.Value ?? DBNull.Value
                                };
                                cmd.Parameters.Add(parameter);
                            }
                        }

                        int affectedRows = cmd.ExecuteNonQuery();
                        return affectedRows > 0;
                    }
                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine("SQL Error: " + sqlEx.Message);
                    // Xem xét việc ghi lại lỗi sqlEx để phân tích thêm
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    // Xem xét việc ghi lại lỗi ex để phân tích thêm
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

    }
}
