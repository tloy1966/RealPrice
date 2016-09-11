using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        public void InsertData()
        {

        }
    }
}
