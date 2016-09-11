using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
namespace RP_Job
{
    class Program
    {
        static string strRPFolder = @"D:\7RealPrice";
        static string testStr = @"D:\7RealPrice\2016S2\A_lvr_land_A.CSV";
        static void Main(string[] args)
        {
            SqlControl.testConn();
            //testGetCityCode();

            Console.ReadLine(); 
        }
        static void testGetCityCode()
        {
            Debug.WriteLine(GetData.GetFiles(strRPFolder,".csv"));
        }
    }
}
