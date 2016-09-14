using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using System.IO;
using NLog;
namespace RP_Job
{
    class Program
    {
        static public Logger logger = NLog.LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            GO();
            //testGetCityCode();
            Console.ReadLine(); 
        }
        static void testGetCityCode()
        {
            var lstPath = GetData.GetFiles(Path.Combine(Directory.GetCurrentDirectory()), ".xls");
            foreach (var file in lstPath)
            {
                logger.Info(file+" test start");
                GetData.ReadXLSAndInsert(file);
                logger.Info(file + " test done");
                logger.Info("------------------------------------------");
            }
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
        }
    }
}
