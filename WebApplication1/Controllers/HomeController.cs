using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {

            var custId = 1234; //ToDo get from session when it is set

            using (MemoryStream mem = new MemoryStream())
            {
                ExcelReportGenerator excel = new ExcelReportGenerator();

                excel.CreateExcelDocNew(mem);

                var reportName = string.Format("MyReport_{0}.xlsx", custId);
                var returnFile = File(mem.ToArray(), System.Net.Mime.MediaTypeNames.Application.Octet, reportName);

                return returnFile;
            }

        }
    }
}