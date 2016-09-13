using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace RP_Job.Model
{
    public class RPModel
    {
        static public DataTable CreateMainData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("City", typeof(string));
            dt.Columns.Add("SellType", typeof(string));
            dt.Columns.Add("DISTRICT", typeof(string));
            dt.Columns.Add("CASE_T", typeof(string));
            dt.Columns.Add("LOCATION", typeof(string));
            dt.Columns.Add("LANDA", typeof(decimal));
            dt.Columns.Add("CASE_F", typeof(string));

            dt.Columns.Add("LANDA_X", typeof(string));
            dt.Columns.Add("LANDA_Z", typeof(string));
            dt.Columns.Add("SDATE", typeof(string));
            dt.Columns.Add("SCNT", typeof(string));

            dt.Columns.Add("SBUILD", typeof(string));
            dt.Columns.Add("TBUILD", typeof(string));
            dt.Columns.Add("BUITYPE", typeof(string));
            dt.Columns.Add("PBUILD", typeof(string));
            dt.Columns.Add("MBUILD", typeof(string));

            dt.Columns.Add("FDATE", typeof(string));
            dt.Columns.Add("FAREA", typeof(decimal));
            dt.Columns.Add("BUILD_R", typeof(int));
            dt.Columns.Add("BUILD_L", typeof(int));
            dt.Columns.Add("BUILD_B", typeof(int));

            dt.Columns.Add("BUILD_P", typeof(string));
            dt.Columns.Add("RULE", typeof(string));
            dt.Columns.Add("BUILD_C", typeof(string));
            dt.Columns.Add("TPRICE", typeof(int));
            dt.Columns.Add("UPRICE", typeof(decimal));

            dt.Columns.Add("PARKTYPE", typeof(string));
            dt.Columns.Add("PAREA", typeof(decimal));
            dt.Columns.Add("PPRICE", typeof(int));
            dt.Columns.Add("RMNOTE", typeof(string));
            dt.Columns.Add("ID2", typeof(string));
            return dt;
        }
        public enum DBCol
        {
            ID = -2,
            City = -1,
            DISTRICT = 0,
            CASE_T = 1,
            LOCATION = 2,
            LANDA = 3,
            CASE_F = 4,

            LANDA_X = 5,
            LANDA_Z = 6,
            SDATE = 7,
            SCNT = 8,
            SBUILD = 9,
            TBUILD = 10,

            BUITYPE = 11,
            PBUILD = 12,
            MBUILD = 13,
            FDATE = 14,
            FAREA = 15,

            BUILD_R = 16,
            BUILD_L = 17,
            BUILD_B = 18,
            BUILD_P = 19,
            RULE = 20,

            //BUILD_C = 21,
            TPRICE = 21,
            UPRICE = 22,
            PARKTYPE = 23,

            PAREA = 24,
            PPRICE = 25,
            RMNOTE = 26,
            ID2 = 27
        }
    }
}
