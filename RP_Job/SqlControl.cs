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
        static public void InsertRow(string strCn, DataTable dt)
        {
            using (SqlConnection cn = new SqlConnection(strCn))
            {
                cn.Open();
                using (SqlBulkCopy bc = new SqlBulkCopy(cn))
                {
                    bc.BulkCopyTimeout = 100;
                    bc.BatchSize = 1;
                    bc.DestinationTableName = "MainData";
                    foreach (var dr in dt.Rows)
                    {
                        try
                        {
                            var tmpDt = dt.Clone();
                            tmpDt.Rows.Add(dr);
                            bc.WriteToServer(tmpDt);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine( ex.Message);
                        }
                        
                    }
                }
            }
            
        }

    }
}
