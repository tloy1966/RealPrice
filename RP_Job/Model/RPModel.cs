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
            dt.Columns.Add("DISTRICT", typeof(string));//0
            dt.Columns.Add("CASE_T", typeof(string));

            dt.Columns.Add("LOCATION", typeof(string));
            dt.Columns.Add("LANDA", typeof(decimal));
            dt.Columns.Add("CASE_F", typeof(string));
            dt.Columns.Add("LANDA_X", typeof(string));
            dt.Columns.Add("LANDA_Z", typeof(string));//6

            dt.Columns.Add("SDATE", typeof(string));
            dt.Columns.Add("SCNT", typeof(string));
            dt.Columns.Add("SBUILD", typeof(string));
            dt.Columns.Add("TBUILD", typeof(string));//10
            dt.Columns.Add("BUITYPE", typeof(string));

            dt.Columns.Add("PBUILD", typeof(string));
            dt.Columns.Add("MBUILD", typeof(string));
            dt.Columns.Add("FDATE", typeof(string));
            dt.Columns.Add("FAREA", typeof(decimal));//15
            dt.Columns.Add("BUILD_R", typeof(int));

            dt.Columns.Add("BUILD_L", typeof(int));
            dt.Columns.Add("BUILD_B", typeof(int));
            dt.Columns.Add("BUILD_P", typeof(string));
            dt.Columns.Add("RULE", typeof(string));//20
                                                   //dt.Columns.Add("BUILD_C", typeof(string));


            dt.Columns.Add("FURNITURE", typeof(bool));
            dt.Columns.Add("TPRICE", typeof(int));
            dt.Columns.Add("UPRICE", typeof(decimal));

            dt.Columns.Add("PARKTYPE", typeof(string));
            dt.Columns.Add("PAREA", typeof(decimal));
            dt.Columns.Add("PPRICE", typeof(int));//25

            dt.Columns.Add("RMNOTE", typeof(string));
            dt.Columns.Add("ID2", typeof(string));
            dt.Columns.Add("isActive", typeof(bool));
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

            FURNITURE = 21,

            TPRICE = 22,
            UPRICE = 23,
            PARKTYPE = 24,

            PAREA = 25,
            PPRICE = 26,
            RMNOTE = 27,
            ID2 = 28,
            isActive = 29

            /*Furniture_C = 21,
            TPRICE_C = 22,
            UPRICE_C = 23,
            PARKTYPE_C = 24,
            PAREA_C = 25,
            PPRICE_C = 26,
            RMNOTE_C = 27,
            ID2_C = 28,
            isActive_C = 29*/
        }
    }
}
