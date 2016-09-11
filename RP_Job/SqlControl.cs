using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
namespace RP_Job
{
    class SqlControl
    {
        public bool testConn()
        {
            using (SqlConnection cn = new SqlConnection())
            {
                try
                {

                }
                catch ()
                {

                }
            }
        }
        public void InsertData()
        {

        }
    }
}
