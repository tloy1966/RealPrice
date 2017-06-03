using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;

namespace RP_Job
{
    class util
    {
        static public List<string> GetFiles(string strFolder, string strFilter)
        {
            DirectoryInfo Dir = new DirectoryInfo(strFolder);
            List<string> lstFiles = new List<string>();
            foreach (var subDir in Dir.GetDirectories())
            {
                foreach (var file in subDir.GetFiles())
                {
                    var t = file.Extension;
                    if (t.Equals(strFilter, StringComparison.OrdinalIgnoreCase))
                    {
                        lstFiles.Add(file.FullName);
                    }
                }
            }
            return lstFiles;
        }



        static public dynamic TypeCheck(string val, string inType, int iLength)
        {
            if (val.Length == 0)
            {
                return DBNull.Value;
            }
            if (inType == "int")
            {
                int iOut;
                if (int.TryParse(val, out iOut))
                {
                    return iOut;
                }
                else
                {
                    return DBNull.Value;
                }
            }
            else if (inType == "decimal")
            {
                decimal dOut;
                if (decimal.TryParse(val, out dOut))
                {
                    return dOut;
                }
                else
                {
                    return DBNull.Value;
                }
            }
            else if (inType == "datetime")
            {
                if (val.Length > 7 || val.Length < 6)
                {
                    Console.WriteLine($"{val} datetime convert error");
                    return DBNull.Value;
                }
                DateTime dt;
                if (DateTime.TryParse(val.Insert(3, "-").Insert(6, "-"), out dt))
                {
                    return dt.AddYears(1911);
                }
                else
                {
                    return DBNull.Value;
                }
            }
            else if (inType == "string")
            {
                if (val.Length > iLength)
                {
                    return val.Substring(0, iLength);
                }
                else
                {
                    return val;
                }
            }
            else if (inType == "letters")
            {
                if (Regex.IsMatch(val, @"^[a-zA-Z0-9]+$"))
                {
                    return val;
                }
                else
                {
                    return DBNull.Value;
                }
            }
            else if (inType == "bool")
            {
                if (val == "有" || val == "Y")
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                return DBNull.Value;
            }
        }
    }
}
