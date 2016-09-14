using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using System.IO;
namespace RP_Job
{
    class Program
    {
        static void Main(string[] args)
        {
            //GO();
            testGetCityCode();
            Console.ReadLine(); 
        }
        static void testGetCityCode()
        {
            var lstPath = GetData.GetFiles(Path.Combine(Directory.GetCurrentDirectory()), ".xls");
            foreach (var file in lstPath)
            {
                GetData.ReadXLSAndInsert(file);
                Console.WriteLine(file);

            }

            Console.ReadLine();
        }

        static void GO()
        {
            var lstPath = GetData.GetFiles(Directory.GetCurrentDirectory(), ".xls");
            foreach (var file in lstPath)
            {
                GetData.ReadXLSAndInsert(file);
                Console.WriteLine(file);
                Thread.Sleep(1000);
            }

            Console.ReadLine();
        }
    }
}
