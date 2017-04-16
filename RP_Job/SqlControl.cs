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
        static public bool testConn(string strCn)
        {
            using (SqlConnection cn = new SqlConnection(strCn))
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
        static public void InsertDtData(string strCn, DataTable dt)
        {
            try
            {
                //Autho.Azure.conn
                using (SqlConnection cn = new SqlConnection(strCn))
                {
                    cn.Open();
                    using (SqlBulkCopy bc = new SqlBulkCopy(cn))
                    {
                        bc.BulkCopyTimeout = 600;
                        bc.BatchSize = 500;
                        bc.DestinationTableName = "MainData";
                        bc.WriteToServer(dt);
                    }
                }
            }
            catch (Exception InsertDtData)
            {
                Program.logger.Error($"Insert error: {InsertDtData.Message}");
                Program.logger.Error($"Insert error: {InsertDtData.StackTrace}");
                Program.logger.Error($"Insert error: {InsertDtData.Source}");
                Console.WriteLine(InsertDtData.Message);
            }
        }
    }
}
