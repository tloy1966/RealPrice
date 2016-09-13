using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
namespace RP_Job
{
    class SqlControl
    {
        static public bool testConn()
        {
            using (SqlConnection cn = new SqlConnection(Autho.Azure.conn))
            {
                try
                {
                    cn.Open();
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    return false;
                }
            }
            return true;
        }
        static public void InsertDtData(DataTable dt)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(Autho.Azure.conn))
                {
                    cn.Open();
                    using (SqlBulkCopy bc = new SqlBulkCopy(cn))
                    {
                        bc.BatchSize = 500;
                        bc.DestinationTableName = "MainData";
                        bc.WriteToServer(dt);
                    }
                }
            }
            catch (Exception InsertDtData)
            {
                Console.WriteLine(InsertDtData.Message);
                Console.ReadLine();
            }
        }
    }
}
