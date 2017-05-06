using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Azure; // Namespace for CloudConfigurationManager
using Microsoft.WindowsAzure.Storage; // Namespace for CloudStorageAccount
using Microsoft.WindowsAzure.Storage.Table; // Namespace for Table storage types
namespace RPStorage.Models
{
    public class RPEntity: TableEntity
    {
        public RPEntity(string city, string ID2)
        {
            this.PartitionKey = city;
            this.RowKey = ID2;
        }
        public RPEntity()
        {
        }
        public string City { get; set; }
        public string SellType { get; set; }
        public string DISTRICT { get; set; }
        public string CASE_T { get; set; }

        public string LOCATION { get; set; }
        public decimal LANDA { get; set; }
        public string CASE_F { get; set; }
        public string LANDA_X { get; set; }
        public string LANDA_Z { get; set; }

        public DateTime SDATE { get; set; }
        public string SCNT { get; set; }
        public string SBUILD { get; set; }
        public string TBUILD { get; set; }
        public string BUITYPE { get; set; }

        public string PBUILD { get; set; }
        public string MBUILD { get; set; }
        public DateTime FDATE { get; set; }
        public decimal FAREA { get; set; }
        public int BUILD_R { get; set; }

        public int BUILD_L { get; set; }
        public int BUILD_B { get; set; }
        public string BUILD_P { get; set; }
        public string RULE { get; set; }
        public bool FURNITURE { get; set; }
        public int TPRICE { get; set; }
        public decimal UPRICE { get; set; }
        public string PARKTYPE { get; set; }
        public decimal PAREA { get; set; }
        public int PPRICE { get; set; }

        public string RMNOTE { get; set; }
        public string ID2 { get; set; }
        //public string isActive { get; set; }
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
