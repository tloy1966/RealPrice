using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace RP_Job
{
    class GetData
    {
        static Regex rgx = new Regex(@"\b[A-Z]{1}_lvr_land_[A-Z]{1}", RegexOptions.IgnoreCase);
        static public void ReadXLSAndInsert(string strXlsPath)
        {
            HSSFWorkbook wk;
            using (FileStream fs = new FileStream(strXlsPath, FileMode.Open, FileAccess.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
            }
            try
            {

                DataTable dt = Model.RPModel.CreateMainData();
                ISheet sheet = null;
                if (wk.GetSheet("不動產買賣") != null)
                {
                    sheet = wk.GetSheet("不動產買賣");//預售屋買賣 不動產租賃
                }
                else if (wk.GetSheet("預售屋買賣") != null)
                {
                    sheet = wk.GetSheet("預售屋買賣");
                }
                else if (wk.GetSheet("不動產租賃") != null)
                {
                    sheet = wk.GetSheet("不動產租賃");
                }
                else
                {
                    return;
                }
                
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    var dr = dt.NewRow();
                    //dr[0] = ""; id
                    var CityCodeAndSellType = GetCityandTypeCode(strXlsPath);
                    dr[1] = CityCodeAndSellType.Item1;
                    dr[2] = CityCodeAndSellType.Item2;
                    var nowRow = sheet.GetRow(i);
                    dr[Model.RPModel.DBCol.CASE_T.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_T)];
                    dr[Model.RPModel.DBCol.DISTRICT.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.DISTRICT)];
                    dr[Model.RPModel.DBCol.LOCATION.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.LOCATION)];
                    dr[Model.RPModel.DBCol.LANDA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].ToString(), "decimal");
                    dr[Model.RPModel.DBCol.CASE_F.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_F)];

                    dr[Model.RPModel.DBCol.LANDA_X.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA_X)];
                    dr[Model.RPModel.DBCol.LANDA_Z.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA_Z)];
                    dr[Model.RPModel.DBCol.SDATE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString(), "datetime");// nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString().Insert(3, "-").Insert(6, "-");//[Model.RPModel.DBCol.SDATE.ToString().Insert(3, "-").Insert(6, "-")]).Date;
                    dr[Model.RPModel.DBCol.SCNT.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.SCNT)];
                    dr[Model.RPModel.DBCol.SBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.SBUILD)];
                    dr[Model.RPModel.DBCol.TBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.TBUILD)];

                    dr[Model.RPModel.DBCol.BUITYPE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUITYPE)];
                    dr[Model.RPModel.DBCol.PBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PBUILD)];
                    dr[Model.RPModel.DBCol.MBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.MBUILD)];
                    dr[Model.RPModel.DBCol.FDATE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.FDATE)].ToString(), "datetime");
                    dr[Model.RPModel.DBCol.FAREA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.FAREA)].ToString(), "decimal");

                    dr[Model.RPModel.DBCol.BUILD_R.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_R)].ToString(), "int");
                    dr[Model.RPModel.DBCol.BUILD_L.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_L)].ToString(), "int");
                    dr[Model.RPModel.DBCol.BUILD_B.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_B)].ToString(), "int");
                    dr[Model.RPModel.DBCol.BUILD_P.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_P)];
                    dr[Model.RPModel.DBCol.RULE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.RULE)];

                    //dr[Model.RPModel.DBCol.BUILD_C.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_C)];
                    dr[Model.RPModel.DBCol.TPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.TPRICE)].ToString(), "int");
                    dr[Model.RPModel.DBCol.UPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.UPRICE)].ToString(), "decimal");
                    //dr[Model.RPModel.DBCol.UPNOTE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.UPNOTE)];
                    dr[Model.RPModel.DBCol.PARKTYPE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PARKTYPE)];

                    dr[Model.RPModel.DBCol.PAREA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PAREA)].ToString(), "decimal");
                    dr[Model.RPModel.DBCol.PPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PPRICE)].ToString(), "decimal");
                    dr[Model.RPModel.DBCol.RMNOTE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.RMNOTE)];
                    dr[Model.RPModel.DBCol.ID2.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.ID2)];
                    dt.Rows.Add(dr);
                }
                if (dt.Rows.Count > 0)
                {
                    SqlControl.InsertDtData(dt);
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                Console.WriteLine(ee.StackTrace);
                Debug.WriteLine(ee.StackTrace);
            }
            
        }
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

        /// <summary>
        /// example : A_lvr_land_A.xls
        /// </summary>
        /// <param name="strPath"></param>
        /// <returns></returns>
        static private Tuple<string, string> GetCityandTypeCode(string strPath)
        {
            //example : A_lvr_land_A.xls
            if (File.Exists(strPath))
            {
                var strName =  Path.GetFileName(strPath);
                if (rgx.IsMatch(strName))
                {
                    return new Tuple<string, string>(strName.Split('_').FirstOrDefault(), strName.Split('.').FirstOrDefault().Split('_').LastOrDefault());
                }
            }
            return null;
        }
        static private dynamic TypeCheck(string val, string inType)
        {
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
            else if (inType=="decimal")
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
                DateTime dt;
                if (val.Length == 0)
                {
                    return DBNull.Value;
                }
                if(DateTime.TryParse(val.Insert(3, "-").Insert(6, "-"), out dt))
                {
                    return dt.AddYears(1911);
                }
                else
                {
                    return DBNull.Value;
                }
            }
            else
            {
                return DBNull.Value;
            }
        }
    }
}
