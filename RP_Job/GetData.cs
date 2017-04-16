using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NLog;
namespace RP_Job
{
    class GetData
    {
        static Logger logger = NLog.LogManager.GetCurrentClassLogger();
        static Regex rgx = new Regex(@"\b[A-Z]{1}_lvr_land_[A-Z]{1}", RegexOptions.IgnoreCase);
        static Regex rgxCheckLetters = new Regex(@"^[a-zA-Z0-9]+$");
        //to do : long mathod
        static public void ReadXLSAndInsert(string strXlsPath, bool isTest, Model.Para.InsertMode mode)
        {
            HSSFWorkbook wk;
            ISheet sheet = null;
            using (FileStream fs = new FileStream(strXlsPath, FileMode.Open, FileAccess.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                logger.Info("Path = "+strXlsPath);
            }
            try
            {
                if (wk.GetSheet("不動產買賣") != null)
                {
                    sheet = wk.GetSheet("不動產買賣");//預售屋買賣 不動產租賃
                    logger.Info("不動產買賣");
                }
                else if (wk.GetSheet("預售屋買賣") != null)
                {
                    sheet = wk.GetSheet("預售屋買賣");
                    logger.Info("不動產買賣");
                }
                else if (wk.GetSheet("不動產租賃") != null)
                {
                    sheet = wk.GetSheet("不動產租賃");
                    logger.Info("不動產買賣");
                }
                else
                {
                    return;
                }
                logger.Info("Processing");
                Console.WriteLine("Processing");

                DataTable dt = Model.RPModel.CreateMainData();
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    
                    var nowRow = sheet.GetRow(i);
                    var t = nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_T)];
                    if (string.IsNullOrEmpty(nowRow.Cells[(int)(Model.RPModel.DBCol.DISTRICT)].ToString()))
                    {
                        continue;
                    }

                    
                    var dr = dt.NewRow();
                    //dr[0] = ""; id
                    var CityCodeAndSellType = GetCityandTypeCode(strXlsPath);
                    dr[1] = CityCodeAndSellType.Item1;
                    dr[2] = CityCodeAndSellType.Item2;
                    dr[Model.RPModel.DBCol.CASE_T.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_T)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.DISTRICT.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.DISTRICT)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.LOCATION.ToString()] = ((string)TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.LOCATION)].ToString(), "string", 400)).Replace("~","-");
                    dr[Model.RPModel.DBCol.LANDA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].ToString(), "decimal", 0);
                    dr[Model.RPModel.DBCol.CASE_F.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_F)].ToString(), "string", 50);

                    dr[Model.RPModel.DBCol.LANDA_X.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA_X)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.LANDA_Z.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA_Z)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.SDATE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString(), "datetime", 0);// nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString().Insert(3, "-").Insert(6, "-");//[Model.RPModel.DBCol.SDATE.ToString().Insert(3, "-").Insert(6, "-")]).Date;
                    dr[Model.RPModel.DBCol.SCNT.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.SCNT)].ToString(), "string", 200);
                    dr[Model.RPModel.DBCol.SBUILD.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.SBUILD)].ToString(), "string", 200);
                    dr[Model.RPModel.DBCol.TBUILD.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.TBUILD)].ToString(), "string", 50);

                    dr[Model.RPModel.DBCol.BUITYPE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUITYPE)].ToString(), "string", 200);
                    dr[Model.RPModel.DBCol.PBUILD.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PBUILD)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.MBUILD.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.MBUILD)].ToString(), "string", 300);
                    dr[Model.RPModel.DBCol.FDATE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.FDATE)].ToString(), "datetime", 0);
                    dr[Model.RPModel.DBCol.FAREA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.FAREA)].ToString(), "decimal", 0);

                    dr[Model.RPModel.DBCol.BUILD_R.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_R)].ToString(), "int", 0);
                    dr[Model.RPModel.DBCol.BUILD_L.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_L)].ToString(), "int", 0);
                    dr[Model.RPModel.DBCol.BUILD_B.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_B)].ToString(), "int", 0);
                    dr[Model.RPModel.DBCol.BUILD_P.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_P)].ToString(), "string", 50);
                    dr[Model.RPModel.DBCol.RULE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.RULE)].ToString(), "string", 50);


                    string tempID2 = "";

                    if (CityCodeAndSellType.Item2 == "A")
                    {
                        /*DB col the same but excel not the same*/
                        dr[Model.RPModel.DBCol.FURNITURE.ToString()] = 0;
                        dr[Model.RPModel.DBCol.TPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.TPRICE) - 1].ToString(), "int", 0);
                        dr[Model.RPModel.DBCol.UPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.UPRICE) - 1].ToString(), "decimal", 0);
                        dr[Model.RPModel.DBCol.PARKTYPE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PARKTYPE) - 1].ToString(), "string", 50);

                        dr[Model.RPModel.DBCol.PAREA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PAREA) - 1].ToString(), "decimal", 0);
                        dr[Model.RPModel.DBCol.PPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PPRICE) - 1].ToString(), "int", 0);
                        dr[Model.RPModel.DBCol.RMNOTE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.RMNOTE) - 1].ToString(), "string", 400);

                        if (string.IsNullOrEmpty(TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2) - 1].ToString(), "string", 400)))
                        {
                            return;
                        }

                        dr[Model.RPModel.DBCol.ID2.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2) - 1].ToString(), "string", 400);
                        var tempID = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2) - 1].ToString(), "letters", 0);
                        if (tempID is DBNull)
                        {
                            dr[Model.RPModel.DBCol.isActive.ToString()] = 0;
                        }
                        else
                        {
                            dr[Model.RPModel.DBCol.isActive.ToString()] = 1;
                            tempID2 = tempID;
                        }
                        
                    }
                    else if (CityCodeAndSellType.Item2 == "C")
                    {
                        dr[Model.RPModel.DBCol.FURNITURE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.FURNITURE)].ToString(), "bool", 0);
                        dr[Model.RPModel.DBCol.TPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.TPRICE)].ToString(), "int", 0);
                        dr[Model.RPModel.DBCol.UPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.UPRICE)].ToString(), "decimal", 0);
                        dr[Model.RPModel.DBCol.PARKTYPE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PARKTYPE)].ToString(), "string", 50);

                        dr[Model.RPModel.DBCol.PAREA.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PAREA)].ToString(), "decimal", 0);
                        dr[Model.RPModel.DBCol.PPRICE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.PPRICE)].ToString(), "int", 0);
                        dr[Model.RPModel.DBCol.RMNOTE.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.RMNOTE)].ToString(), "string", 400);

                        dr[Model.RPModel.DBCol.ID2.ToString()] = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2)].ToString(), "string", 400);
                        if (string.IsNullOrEmpty(TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2)].ToString(), "string", 400)))
                        {
                            return;
                        }
                        var tempID = TypeCheck(nowRow.Cells[(int)(Model.RPModel.DBCol.ID2)].ToString(), "letters", 0);
                        if (tempID is DBNull)
                        {
                            dr[Model.RPModel.DBCol.isActive.ToString()] = 0;
                        }
                        else
                        {
                            dr[Model.RPModel.DBCol.isActive.ToString()] = 1;
                            tempID2 = tempID;
                        }
                    }

                    if (string.IsNullOrEmpty(tempID2))
                    {
                        Console.WriteLine("ID null");
                        continue;
                    }
                    dt.Rows.Add(dr);
                }
                if (dt.Rows.Count > 0 && mode== Model.Para.InsertMode.Bulk)
                {
                    Console.WriteLine($"InsertDtData{strXlsPath}, rows count = {dt.Rows.Count}");
                    Program.logger.Info($"Insert: {strXlsPath}, count = {sheet.LastRowNum}");
                    try
                    {
                        SqlControl.InsertDtData(Autho.Azure.getConnect(isTest), dt);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine("Path: "+strXlsPath);
                        SqlControl.InsertRow(Autho.Azure.getConnect(isTest), dt);
                    }
                }
                else
                {
                    Program.logger.Info($"No data: {strXlsPath}, count = {sheet.LastRowNum}");
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                Console.WriteLine(ee.StackTrace);
                Debug.WriteLine(ee.StackTrace);
                Program.logger.Error(ee.Message);
                Program.logger.Error(ee.StackTrace);
                Program.logger.Error(ee.Source);
            }
            finally
            {
                sheet = null;
                wk = null;
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
        static private dynamic TypeCheck(string val, string inType, int iLength)
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
