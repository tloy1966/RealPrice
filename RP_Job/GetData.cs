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
                int iLastCol = 28;
                var sheet = wk.GetSheet("不動產買賣");
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
                    dr[Model.RPModel.DBCol.CASE_F.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.CASE_F)];
                    dr[Model.RPModel.DBCol.LOCATION.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.LOCATION)];
                    var tt = nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].StringCellValue;
                    Debug.WriteLine(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].StringCellValue);
                    dr[Model.RPModel.DBCol.LANDA.ToString()] =//if(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].ToString()!="") decimal.Parse(nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA)].StringCellValue);

                    dr[Model.RPModel.DBCol.LANDA_Z.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.LANDA_Z)];
                    dr[Model.RPModel.DBCol.SDATE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString().Insert(3, "-").Insert(6, "-");//[Model.RPModel.DBCol.SDATE.ToString().Insert(3, "-").Insert(6, "-")]).Date;
                    dr[Model.RPModel.DBCol.SCNT.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)];
                    dr[Model.RPModel.DBCol.SBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.SBUILD)];
                    dr[Model.RPModel.DBCol.TBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.TBUILD)];

                    dr[Model.RPModel.DBCol.BUITYPE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUITYPE)];
                    dr[Model.RPModel.DBCol.PBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PBUILD)];
                    dr[Model.RPModel.DBCol.MBUILD.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.MBUILD)];
                    dr[Model.RPModel.DBCol.FDATE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.FDATE)].ToString()==""?"": nowRow.Cells[(int)(Model.RPModel.DBCol.FDATE)].ToString().Insert(3, "-").Insert(6, "-");
                    dr[Model.RPModel.DBCol.FAREA.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.FAREA)];

                    dr[Model.RPModel.DBCol.BUILD_R.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_R)];
                    dr[Model.RPModel.DBCol.BUILD_L.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_L)];
                    dr[Model.RPModel.DBCol.BUILD_B.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_B)];
                    dr[Model.RPModel.DBCol.BUILD_P.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_P)];
                    dr[Model.RPModel.DBCol.RULE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.RULE)];

                    dr[Model.RPModel.DBCol.BUILD_C.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.BUILD_C)];
                    dr[Model.RPModel.DBCol.TPRICE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.TPRICE)];
                    dr[Model.RPModel.DBCol.UPRICE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.UPRICE)];
                    dr[Model.RPModel.DBCol.UPNOTE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.UPNOTE)];
                    dr[Model.RPModel.DBCol.PARKTYPE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PARKTYPE)];

                    dr[Model.RPModel.DBCol.PAREA.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PAREA)];
                    dr[Model.RPModel.DBCol.PPRICE.ToString()] = nowRow.Cells[(int)(Model.RPModel.DBCol.PPRICE)];
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
                Debug.WriteLine(ee.Message);
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

    }
}
