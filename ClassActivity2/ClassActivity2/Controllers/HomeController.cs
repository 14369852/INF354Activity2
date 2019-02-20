using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClassActivity2.Models;
//using ClassActivity2.Reports;
using System.Data;
//using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using ClassActivity2.ViewModel;

namespace ClassActivity2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Report()
        {
            Report rep = new Report();
            rep.Employee = GetEmployees(0);


            //note
            rep.DFrom = new DateTime();
            rep.DTo = new DateTime();

            return View(rep);
        }

        private SelectList GetEmployees(int selected)
        {
            using (HardwareDBEntities db = new HardwareDBEntities())
            {
                db.Configuration.ProxyCreationEnabled = false;
                //var invoices = db.lginvoices.Include("lgemployee").ToList();
                /*if (selected == 0)
                    return new SelectList(invoices, "employee_id", "emp_fname");
                else
                    return new SelectList(invoices, "employee_id", "emp_fname", selected)
                 */
                var invoices = db.lgemployees.Select(x => new SelectListItem
                {
                    Value = x.emp_num.ToString(),
                    Text = x.emp_fname.ToString()
                }).ToList();
                if (selected == 0)
                    return new SelectList(invoices, "Value", "Text");
                else
                    return new SelectList(invoices, "Value", "Text", selected);
            }
        }

        [HttpPost]
        public ActionResult Advanced(Report rep)
        {
            using (HardwareDBEntities db = new HardwareDBEntities())
            {
                db.Configuration.ProxyCreationEnabled = false;

                //Retrieve a list of vendors so that it can be used to populate the dropdown on the View
                //The ID of the currently selected item is passed through so that the returned list has that item preselected
                rep.Employee = GetEmployees(rep.SelectedEmployeeID);

                //Get the full details of the selected vendor so that it can be displayed on the view
                rep.Employee = db.lgemployees.Where(x=>x.emp_num == rep.SelectedEmployeeID).FirstOrDefault();

                //Get all supplier orders that adheres to the entered criteria
                //For each of the results, load data into a new ReportRecord object
                var list = db.lginvoices.Include("lgemployees").Where(pp => pp.employee_id == db.lgemployees.emp_id && pp.OrderDate >= rep.DateFrom && pp.OrderDate <= vm.DateTo).ToList().Select(rr => new ReportRecord
                {
                    r
                    = rr.OrderDate.ToString("dd-MMM-yyyy"),
                     = Convert.ToDouble(rr.TotalDue),
                     Customer= rr.Customer.Name,
                       EmployeeID = db.People.Where(pp => pp.BusinessEntityID == rr.EmployeeID).Select(x => x.FirstName + " " + x.LastName).FirstOrDefault(),
                     EmployeeID= rr.EmployeeID
                });

                //Load the list of ReportRecords returned by the above query into a new list grouped by Shipment Method
                rep.results = list.GroupBy(g => g.ShipMethod).ToList();

                //Load the list of ReportRecords returned by the above query into a new dictionary grouped by Employee
                //This will be used to generate the chart on the View through the MicroSoft Charts helper
                rep.chartData = list.GroupBy(g => g.Employee).ToDictionary(g => g.Key, g => g.Sum(v => v.Amount));

                //Store the chartData dictionary in temporary data so that it can be accessed by the EmployeeOrdersChart action resonsible for generating the chart
                TempData["chartData"] = rep.chartData;
                TempData["records"] = list.ToList();
                TempData["employee"] = rep.employee;
                return View(rep);
            }

        }

        //This action returns the EmployeeOrdersChart partial view, which is used to generate a chart for the Advanced report
        public ActionResult EmployeeOrdersChart()
        {
            //Load the chartData from temporary memory
            var data = TempData["chartData"];

            //Return the EmployeeOrdersChart temporary view, pass through the required chartData
            return View(TempData["chartData"]);
        }

        //This action will be called from the Export report functions
        //The purpose of this action is to convert data formatted for the on screen report to a format that is appropriate for the Crystal Report
        private VendorSalesModel GetAdvancedDataSet()
        {
            VendorSalesModel data = new VendorSalesModel();

            data.Vendor.Clear();
            data.VendorSales.Rows.Clear();

            //Add table (with only one record) to dataset for general vendor details to be shown on Crystal Report
            DataRow vrow = data.Employee.NewRow();
            lgemployee employee = (lgemployee)TempData["employee"];
            vrow["ID"] = employee.BusinessEntityID;
            vrow["Name"] = employee.Name;
            vrow["Credit"] = employee.CreditRating;
            vrow["Preferred"] = employee.PreferredVendorStatus;
            data.Vendor.Rows.Add(vrow);

            //Add table to dataset for general vendor details to be shown on Crystal Report
            foreach (var item in (IEnumerable<ReportRecord>)TempData["records"])
            {
                DataRow row = data.VendorSales.NewRow();
                row["OrderDate"] = item.OrderDate;
                row["Amount"] = item.Amount;
                row["ShipMethod"] = item.ShipMethod;
                row["Employee"] = item.EmployeeID;
                row["VendorID"] = item.VendorID;
                data.VendorSales.Rows.Add(row);
            }

            //Reset TempData so that it can be accessed in upcoming calls to controller actions
            TempData["chartData"] = TempData["chartData"];
            TempData["records"] = TempData["records"];
            TempData["vendor"] = TempData["vendor"];
            return data;
        }



        public ActionResult ExportAdvancedPDF()
        {
            ReportDocument report = new ReportDocument();
            report.Load(Path.Combine(Server.MapPath("~/Reports/VendorSales.rpt")));
            report.SetDataSource(GetAdvancedDataSet());
            Response.Buffer = false;
            Response.ClearContent();
            Response.ClearHeaders();
            Stream stream = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            stream.Seek(0, SeekOrigin.Begin);
            return File(stream, "application/pdf", "VendorSales.pdf");
        }

        public ActionResult ExportAdvancedWord()
        {
            ReportDocument report = new ReportDocument();
            report.Load(Path.Combine(Server.MapPath("~/Reports/VendorSales.rpt")));
            report.SetDataSource(GetAdvancedDataSet());
            Response.Buffer = false;
            Response.ClearContent();
            Response.ClearHeaders();
            Stream stream = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.WordForWindows);
            stream.Seek(0, SeekOrigin.Begin);
            return File(stream, "application/msword", "VendorSales.doc");
        }
    }

    public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}