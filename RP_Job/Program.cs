using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
namespace RP_Job
{
    class Program
    {
        static string strRPFolder = @"D:\7RealPrice";
        static string testStr = @"D:\7RealPrice\2016S2\A_lvr_land_A.xls";
        static void Main(string[] args)
        {
            var lstPath = GetData.GetFiles(strRPFolder,".xls");
            foreach (var file in lstPath)
            {
                GetData.ReadXLSAndInsert(file);
                Console.WriteLine(file);
                Thread.Sleep(5000);
            }
            
            Console.ReadLine(); 
        }
        static void testGetCityCode()
        {
            Debug.WriteLine(GetData.GetFiles(Autho.LocalData.strCSV_Folder, ".xls"));
        }
    }
}
