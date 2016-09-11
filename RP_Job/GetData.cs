using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
namespace RP_Job
{
    class GetData
    {
        static Regex rgx = new Regex(@"\b[A-Z]{1}_lvr_land_[A-Z]{1}", RegexOptions.IgnoreCase);

        static public List<string> GetFiles(string strFolder, string Filter)
        {
            DirectoryInfo Dir = new DirectoryInfo(strFolder);
            List<string> lstFiles = new List<string>();
            foreach (var subDir in Dir.GetDirectories())
            {
                foreach (var file in subDir.GetFiles())
                {
                    var t = file.Extension;
                    if (t.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        lstFiles.Add(file.FullName);
                    }
                }
            }
            return lstFiles;
        }
        static public string GetCityCode(string strPath)
        {
            if (File.Exists(strPath))
            {
                var strName =  Path.GetFileName(strPath);
                if (rgx.IsMatch(strName))
                {
                    return strName.Split('_').FirstOrDefault();
                }
            }
            

            return "";
        }
    }
}
