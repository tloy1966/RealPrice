using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using RP_Job;
namespace RPStorage
{
    class Program
    {

        /// <summary>
        /// Seems like a cheap way to storage data, let's check performance
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var _util = new util();
            //_util.Create();
            TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks);
            _util.GetInstance();
            TimeSpan ts2 = new TimeSpan(DateTime.Now.Ticks);
            Console.WriteLine();
            Console.WriteLine((ts2-ts1).TotalSeconds);
            Console.ReadLine();
            /*string path = @"C:\TestData\2Doing";

            var lstPath = GetData.GetFiles(path, ".xls").OrderBy(o => o);
            foreach (var file in lstPath)
            {
                var sheet = RP_Job.GetData.RealXls(file);
                var lst = util.ReadExcelFileToList(sheet);
                util.Insert(lst);
            }
            */
        }
    }
}
