using ClassActivity2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ClassActivity2.ViewModel
{
    public class Report
    {
        public IEnumerable<SelectListItem> Employee { get; set; }
        public int SelectedEmployeeID { get; set; }
        public DateTime DFrom { get; set; }
        public DateTime DTo { get; set; }


        public lgemployee employee {get; set;}
        public List <IGrouping<string, ReportRecord>> results { get; set; }
        public Dictionary <string,double> chartData { get; set; }
    }

    public class ReportRecord
    {
        public string Date { get; set; }
        public double Amount { get; set; }
        public string InvoiceNumber { get; set; }
        public string Customer { get; set; }
        public string EmployeeID { get; set; }
    }
}