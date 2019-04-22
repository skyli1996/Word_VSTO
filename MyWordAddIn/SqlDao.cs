using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWordAddIn
{
    public class SqlDao
    {
        ////连接数据库的步骤：
        ////1.创建连接字符串
        ////Data Source=服务器名;
        ////Initial Catalog=数据库名;
        ////Integrated Security=True;声明验证方式
        ////用户名、密码方式
        static string MySqlCon = "Data Source=DESKTOP-0MKMHN0\\SQLEXPRESS;Initial Catalog=Sky;Integrated Security=True";

        public DataTable ExecuteQuery(string sqlStr)
        {
            using (SqlConnection con = new SqlConnection(@MySqlCon))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = sqlStr;
                DataTable dt = new DataTable();
                SqlDataAdapter msda;
                msda = new SqlDataAdapter(cmd);
                msda.Fill(dt);
                con.Close();
                return dt;
            }
        }

        public int ExecuteUpdate(string sqlStr)
        {
            using (SqlConnection con = new SqlConnection(@MySqlCon))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = sqlStr;
                int iud = 0;
                iud = cmd.ExecuteNonQuery();
                con.Close();
                return iud;
            }
        }
    }
}
