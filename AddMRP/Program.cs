using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
namespace AddMRP
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Path.Combine(Environment.CurrentDirectory, "Assets", "MRT_W_GEO.txt");
            if (!File.Exists(path))
            {
                return;
            }
            var lines = File.ReadAllLines(path);
            using (SqlConnection cn = new SqlConnection(RP_Job.Autho.Azure.getConnect(false)))
            {
                cn.Open();

                foreach (var line in lines)
                {
                    var geo = line.Split('|');
                    string sql = $"INSERT INTO [dbo].[MRTGeo] ([MRT],[lat],[lng],[place_ID],[location]) values ('{geo[0]}',{geo[1]},{geo[2]},'{geo[3]}','{geo[4]}')";
                    SqlCommand cmd = new SqlCommand(sql, cn);
                    try
                    {

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {

                        Console.WriteLine(e.Message);
                    }
                }
            }
            Console.ReadLine();
        }
    }
}
