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
        static readonly bool isTest = false;

        static public Logger logger = NLog.LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            GetDataFromFile();
            //testGetCityCode();
            Console.ReadLine(); 
        }
        static void testGetCityCode()
        {
            var lstPath = GetData.GetFiles(Path.Combine(Directory.GetCurrentDirectory()), ".xls").OrderByDescending(o => o);
            foreach (var file in lstPath)
            {
                logger.Info(file+" test start");
                GetData.ReadXLSAndInsert(file,true, Model.Para.InsertMode.OneByOne);
                logger.Info(file + " test done");
                logger.Info("------------------------------------------");
            }
        }

        //Type B not in ...............
        static void GetDataFromFile()
        {
            var lstPath = GetData.GetFiles(Autho.LocalData.strXLSFolder_2, ".xls").OrderBy(o => o);
            foreach (var file in lstPath)
            {
                GetData.ReadXLSAndInsert(file, isTest, Model.Para.InsertMode.OneByOne);
                Console.WriteLine(file);
                Thread.Sleep(1000);
            }
        }

        static void GetDataFromAPI()
        {

        }
    }
}
