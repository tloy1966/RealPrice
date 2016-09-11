using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RP_Job.Model
{
    public class RPModel
    {
        public enum DBCol
        {
            ID = -1,
            City = 0,
            CASE_T = 1,
            DISTRICT = 2,
            CASE_F = 3,
            LOCATION = 4,
            LANDA = 5,

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

            BUILD_C = 21,
            TPRICE = 22,
            UPRICE = 23,
            UPNOTE = 24,
            PARKTYPE = 25,

            PAREA = 26,
            PPRICE = 27,
            RMNOTE = 28
        }
    }
}
