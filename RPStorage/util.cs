using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.Azure; // Namespace for CloudConfigurationManager
using Microsoft.WindowsAzure.Storage; // Namespace for CloudStorageAccount
using Microsoft.WindowsAzure.Storage.Table; // Namespace for Table storage types
using NLog;
using System.Text.RegularExpressions;
using System.IO;
using Newtonsoft.Json;
namespace RPStorage
{
    public class util
    {
        static Logger logger = NLog.LogManager.GetCurrentClassLogger();
        static Regex rgxFileName = new Regex(@"\b[A-Z]{1}_lvr_land_[A]{1}", RegexOptions.IgnoreCase);

        /// <summary>
        /// Table = City
        /// Partition key = Selltype_Pbuild_BuiType (A_住家用_住宅大樓)
        /// Row key = id2
        /// </summary>
        static public void Insert(List<Models.RPEntity> lst)
        {
            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
                CloudConfigurationManager.GetSetting("StorageConnectionString"));
                // Create the table client.
                CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
                tableClient.Credentials.TransformUri(new Uri("https://rpscore.table.core.windows.net/?sv=2016-05-31&ss=bfqt&srt=sco&sp=rwdlacup&se=2017-05-18T21:29:09Z&st=2017-05-03T13:29:09Z&spr=https,http&sig=aOinNnNV3sq0Za71%2BF4V%2BMHNe9NlpZV0ejFf3lQuwds%3D"));
                foreach (var nosql in lst)
                {
                    CloudTable table = tableClient.GetTableReference($"city{nosql.City}");

                    try
                    {
                        table.CreateIfNotExists();
                        TableOperation insertOperation = TableOperation.Insert(nosql);
                        table.Execute(insertOperation);
                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss")+$" {nosql.ID2}  OK");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                        logger.Error(nosql.ID2);
                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss")+" : "+ex.Message);
                    }
                    
                }
            }
            catch(Exception ex)
            {
                logger.Info(ex);
            }
        }

        public void Create()
        {
            // Parse the connection string and return a reference to the storage account.
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
                CloudConfigurationManager.GetSetting("StorageConnectionString"));
            // Create the table client.
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            
            // Retrieve a reference to the table.
            CloudTable table = tableClient.GetTableReference("people");
            var tabless = tableClient.ListTables();
            try
            {
                // Create the table if it doesn't exist.
                table.CreateIfNotExists();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            // Execute the insert operation.
            for (int i = 1; i < 100000; i++)
            {

                CustomerEntity customer1 = new CustomerEntity("Harp4", "Walter");
                customer1.Email = "Walter@contoso.com"+i.ToString();
                customer1.PhoneNumber = "425-555-0101";

                // Create the TableOperation object that inserts the customer entity.
                TableOperation insertOperation = TableOperation.Insert(customer1);
                table.Execute(insertOperation);
                Console.WriteLine(i.ToString());
            }

        }

        public void GetInstance()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
                    CloudConfigurationManager.GetSetting("StorageConnectionString"));
            // Create the table client.
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

            // Retrieve a reference to the table.
            CloudTable table = tableClient.GetTableReference("cityB");

            var query = new TableQuery<Models.RPEntity>();//.Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, "Smith"));
                                                          // Print the fields for each customer.
            for (int i = 0; i < 100; i++)
            {
            TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks);
            var q1 = table.ExecuteQuery(query);
                Console.WriteLine("B  " + q1.Count());

                TimeSpan ts2 = new TimeSpan(DateTime.Now.Ticks);
                Console.WriteLine((ts2 - ts1).TotalSeconds);

                Console.WriteLine();

            }//TimeSpan ts1 = DateTime.Now;


            /*table = tableClient.GetTableReference("cityB");
            q1 = table.ExecuteQuery(query);
            Console.WriteLine("B  " +q1.Count());

            table = tableClient.GetTableReference("cityB");
            q1 = table.ExecuteQuery(query);
            Console.WriteLine("C  " + q1.Count());
            */

        }

        static public List<Models.RPEntity> ReadExcelFileToList(Tuple<ISheet, string> iSheet)
        {
            List<Models.RPEntity> lst = new List<Models.RPEntity>();
            var sheet = iSheet.Item1;
            var strXlsPath = iSheet.Item2;
            var CityCodeAndSellType = GetCityandTypeCode(strXlsPath);
            for (int i = 2; i <= sheet.LastRowNum; i++)
            {
                try
                {
                    var nowRow = sheet.GetRow(i);
                    var t = nowRow.Cells[(int)(Models.DBCol.CASE_T)];
                    if (string.IsNullOrEmpty(nowRow.Cells[(int)(Models.DBCol.DISTRICT)].ToString()))
                    {
                        continue;
                    }
                    if (CityCodeAndSellType == null)
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(TypeCheck(nowRow.Cells[(int)(Models.DBCol.ID2) - 1].ToString(), "string", 400)))
                    {
                        continue; ;
                    }
                    var tempID = TypeCheck(nowRow.Cells[(int)(Models.DBCol.ID2) - 1].ToString(), "letters", 0);
                    if (tempID is DBNull)
                    {
                        continue;
                    }
                    var tempID2 = TypeCheck(nowRow.Cells[(int)(Models.DBCol.ID2) - 1].ToString(), "string", 400);
                    Models.RPEntity nosql = new Models.RPEntity(CityCodeAndSellType.Item1, tempID2);
                    nosql.ID2 = TypeCheck(nowRow.Cells[(int)(Models.DBCol.ID2) - 1].ToString(), "string", 400);
                    nosql.City = CityCodeAndSellType.Item1;
                    nosql.SellType = CityCodeAndSellType.Item2;

                    nosql.CASE_T = TypeCheck(nowRow.Cells[(int)(Models.DBCol.CASE_T)].ToString(), "string", 50);
                    nosql.DISTRICT = TypeCheck(nowRow.Cells[(int)(Models.DBCol.DISTRICT)].ToString(), "string", 50);
                    nosql.LOCATION = ((string)TypeCheck(nowRow.Cells[(int)(Models.DBCol.LOCATION)].ToString(), "string", 400)).Replace("~", "-");
                    nosql.LANDA = TypeCheck(nowRow.Cells[(int)(Models.DBCol.LANDA)].ToString(), "decimal", 0);
                    nosql.CASE_F = TypeCheck(nowRow.Cells[(int)(Models.DBCol.CASE_F)].ToString(), "string", 50);

                    var temp2 = nowRow.Cells[(int)(Models.DBCol.LANDA_X)];
                    nosql.LANDA_X = TypeCheck(nowRow.Cells[(int)(Models.DBCol.LANDA_X)].ToString(), "string", 50);
                    nosql.LANDA_Z = TypeCheck(nowRow.Cells[(int)(Models.DBCol.LANDA_Z)].ToString(), "string", 50);
                    nosql.SDATE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.SDATE)].ToString(), "datetime", 0);// nowRow.Cells[(int)(Model.RPModel.DBCol.SDATE)].ToString().Insert(3, "-").Insert(6, "-");//[Model.RPModel.DBCol.SDATE.ToString().Insert(3, "-").Insert(6, "-")]).Date;
                    nosql.SCNT = TypeCheck(nowRow.Cells[(int)(Models.DBCol.SCNT)].ToString(), "string", 200);
                    nosql.SBUILD = TypeCheck(nowRow.Cells[(int)(Models.DBCol.SBUILD)].ToString(), "string", 200);
                    nosql.TBUILD = TypeCheck(nowRow.Cells[(int)(Models.DBCol.TBUILD)].ToString(), "string", 50);

                    nosql.BUITYPE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.BUITYPE)].ToString(), "string", 200);
                    nosql.PBUILD = TypeCheck(nowRow.Cells[(int)(Models.DBCol.PBUILD)].ToString(), "string", 50);
                    nosql.MBUILD = TypeCheck(nowRow.Cells[(int)(Models.DBCol.MBUILD)].ToString(), "string", 300);
                    nosql.FDATE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.FDATE)].ToString(), "datetime", 0);
                    nosql.FAREA = TypeCheck(nowRow.Cells[(int)(Models.DBCol.FAREA)].ToString(), "decimal", 0);

                    nosql.BUILD_R = TypeCheck(nowRow.Cells[(int)(Models.DBCol.BUILD_R)].ToString(), "int", 0);
                    nosql.BUILD_L = TypeCheck(nowRow.Cells[(int)(Models.DBCol.BUILD_L)].ToString(), "int", 0);
                    nosql.BUILD_B = TypeCheck(nowRow.Cells[(int)(Models.DBCol.BUILD_B)].ToString(), "int", 0);
                    nosql.BUILD_P = TypeCheck(nowRow.Cells[(int)(Models.DBCol.BUILD_P)].ToString(), "string", 50);
                    nosql.RULE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.RULE)].ToString(), "string", 50);


                    /*DB col the same but excel not the same*/
                    nosql.FURNITURE = false;
                    nosql.TPRICE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.TPRICE) - 1].ToString(), "int", 0);
                    nosql.UPRICE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.UPRICE) - 1].ToString(), "decimal", 0);
                    nosql.PARKTYPE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.PARKTYPE) - 1].ToString(), "string", 50);

                    nosql.PAREA = TypeCheck(nowRow.Cells[(int)(Models.DBCol.PAREA) - 1].ToString(), "decimal", 0);
                    nosql.PPRICE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.PPRICE) - 1].ToString(), "int", 0);
                    nosql.RMNOTE = TypeCheck(nowRow.Cells[(int)(Models.DBCol.RMNOTE) - 1].ToString(), "string", 400);


                    nosql.ID2 = TypeCheck(nowRow.Cells[(int)(Models.DBCol.ID2) - 1].ToString(), "string", 400);

                    lst.Add(nosql);
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
                
            }
            return lst;
        }
        static public dynamic TypeCheck(string val, string inType, int iLength)
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
            else if (inType == "decimal")
            {
                decimal dOut;
                if (decimal.TryParse(val, out dOut))
                {
                    return dOut;
                }
                else
                {
                    return -1;
                }
            }
            else if (inType == "datetime")
            {
                if (val.Length > 7 || val.Length < 6)
                {
                    return new DateTime(1900, 1, 1);
                }
                DateTime dt;
                if (DateTime.TryParse(val.Insert(3, "-").Insert(6, "-"), out dt))
                {
                    return dt.AddYears(1911);
                }
                else
                {
                    return new DateTime(1900, 1, 1);
                }
            }
            else if (inType == "string")
            {
                if (val.Length <= 0)
                {
                    return "";
                }
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

        static private Tuple<string, string> GetCityandTypeCode(string strPath)
        {
            //example : A_lvr_land_A.xls
            if (File.Exists(strPath))
            {
                var strName = Path.GetFileName(strPath);
                if (rgxFileName.IsMatch(strName))
                {
                    return new Tuple<string, string>(strName.Split('_').FirstOrDefault(), strName.Split('.').FirstOrDefault().Split('_').LastOrDefault());
                }
            }
            return null;
        }

    }



    public class CustomerEntity : TableEntity
    {
        public CustomerEntity(string lastName, string firstName)
        {
            this.PartitionKey = lastName;
            this.RowKey = firstName;
        }

        public CustomerEntity() { }

        public string Email { get; set; }

        public string PhoneNumber { get; set; }
    }
}
