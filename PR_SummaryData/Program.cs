using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PR_SummaryData
{
    class Program
    {
        
        static void Main(string[] args)
        {
            summary();
        }
        static void summary()
        {
            DataClasses1DataContext dc = new DataClasses1DataContext();
            dc.CommandTimeout = 20000;
            var summary = from x in dc.MainData
                          group x by new {
                              x.City,
                              x.SellType,
                              x.DISTRICT,
                              x.CASE_T,
                              x.LOCATION
                          }
                          into g
                          select new {
                              Ciry = g.Key.City,
                              SellType = g.Key.SellType,
                              DISTRICT = g.Key.DISTRICT,
                              CASE_T = g.Key.CASE_T,
                              LOCATION = g.Key.LOCATION,
                              Avg = (double)g.Average(t=>t.UPRICE)
                          };
            //select  city,selltype,[DISTRICT],case_t, location,avg(uprice)
            // from dbo.maindata group bycity, selltype, [DISTRICT], case_t, location

            foreach (var x in summary)
            {
                var newSumm = new SummaryData
                {
                    City = x.Ciry,
                    SellType = x.SellType,
                    DISTRICT = x.DISTRICT,
                    CASE_T = x.CASE_T,
                    LOCATION = x.LOCATION,
                    Avg = x.Avg
                };
                dc.SummaryData.InsertOnSubmit(newSumm);
                dc.SubmitChanges();
            }                  

        }
    }
}
