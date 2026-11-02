using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OYA_Web_Application.Hubs;
using OYA_Web_Application.Models;
using OYA_Web_Application.Models.Data.Dashboard;
using OYA_Web_Application.Models.Data.QBSync;
using OYA_Web_Application.Models.ViewModel;
using OYA_Web_Application.Services.ReportServices;
using System.Diagnostics;

namespace OYA_Web_Application.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;
        string connectionString = "";
        private readonly IWebHostEnvironment _webHostEnvironment;

        public HomeController(IConfiguration configuration, ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            _logger = logger;
            _configuration = configuration;
            connectionString = _configuration.GetConnectionString("DefaultConnection");
        }

        public IActionResult Dashboard()
        {
            SalesBifurcation salesDashboard = new SalesBifurcation(_configuration);

            DashboardVM dashboard = new DashboardVM
            {
                dashboardSales = salesDashboard.GetOverallDashboardSales(),
                customerSales = salesDashboard.GetTopCustomerSales(),
                oyaMonthlySales = salesDashboard.GetOYAMonthlySales(),
                oyaYearlySales = salesDashboard.GetOYAYearlySales(),
                pinoMonthlySales = salesDashboard.GetPinoMonthlySales(),
                pinoYearlySales = salesDashboard.GetPinoYearlySales(),
            };
            return View(dashboard);
        }

        public IActionResult Logs()
        {
            return View();
        }

        public IActionResult Users()
        {
            return View();
        }

        public IActionResult Reports()
        {
            SalesBifurcation salesDashboard = new SalesBifurcation(_configuration);
            DailyReports dailyReports = new DailyReports(_configuration);
            ViewBag.Contractors = dailyReports.GetContractors();

            return View(salesDashboard.GetOverallDashboardSales());
        }
        public IActionResult Sync()
        {
            SycnServices sycnServices = new SycnServices(_configuration);
            return View(sycnServices.GetAllSyncDetails());
        }
        [HttpGet]
        public async Task<IActionResult> SyncNow(CancellationToken cancellationToken)
        {
            try
            {
                await new PubNub(_configuration).SendMessageAsync(
                    new Dictionary<string, string> { { "msg", "start sync" } }, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Could not send the QuickBooks sync command");
                TempData["SyncPublishError"] = "The sync command could not be sent. Check the server's PubNub configuration and logs.";
            }
            return RedirectToAction("Sync");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
        public async Task<IActionResult> SyncSubscriptionToken(CancellationToken cancellationToken)
        {
            const string sessionKey = "PubNub.BrowserUserId";
            var userId = HttpContext.Session.GetString(sessionKey);
            if (string.IsNullOrWhiteSpace(userId))
            {
                userId = "OYA_WEB_BROWSER_" + Guid.NewGuid().ToString("N");
                HttpContext.Session.SetString(sessionKey, userId);
            }
            try
            {
                var client = new PubNub(_configuration);
                var token = await client.CreateReadTokenAsync(userId, cancellationToken);
                return Json(new { token, userId, subscribeKey = client.SubscribeKey,
                    channel = PubNub.ChannelName, renewAfterSeconds = 600 });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Could not authorize PubNub sync status updates");
                return StatusCode(503, new { error = "Live sync updates are unavailable. Check the server's PubNub configuration and logs." });
            }
        }
        [HttpPost]
        public IActionResult DownloadReport(IFormCollection frm, string reportType, string fromDate, string toDate)
        {
            // Get the file path from the configuration
            string filePath = "";
            string fileName = "";
            string webRootPath = _webHostEnvironment.WebRootPath;
            reportType = string.IsNullOrEmpty(reportType) ? Convert.ToString(frm["reportType"]) : reportType;
            DateTime startVal;
            DateTime endVal;
            if (string.IsNullOrEmpty(fromDate))
            {
                startVal = Convert.ToDateTime(DateTime.Now.Month + "/01/" + DateTime.Now.Year);
            }
            else
            {
                startVal = Convert.ToDateTime(fromDate);
            }
            if (string.IsNullOrEmpty(toDate))
            {
                endVal = Convert.ToDateTime(DateTime.Now.Month + "/01/" + DateTime.Now.Year);
            }
            else
            {
                endVal = Convert.ToDateTime(toDate);
            }
            switch (reportType)
            {
                case "HeliasDailyReport":
                    DailyReports dr = new DailyReports(_configuration);
                    //filePath = _configuration["FilePaths:HeliasDailyReport"];
                    fileName = "Helias_Sales_Report";
                    return File(dr.DownloadExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");

                case "Territory_Report":
                    {
                        filePath = _configuration["FilePaths:TerritoryReport"];
                        fileName = "Helias Sales Report - Territory";
                    }
                    break;
                case "AR_Aging_Detail_Oya":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary OYA";
                        return File(ar.DownloadExcel("OYA", "Details"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Detail_Pino":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary PINO";
                        return File(ar.DownloadExcel("PINO", "Details"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Summary_Oya":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary OYA";
                        return File(ar.DownloadExcel("OYA"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Summary_Pino":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary PINO";
                        return File(ar.DownloadExcel("PINO"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                default:
                    {
                        var contractor = reportType.Split('_')[1];
                        CommissionReport cr = new CommissionReport(_configuration);
                        fileName = "Commission Report - " + contractor;
                        return File(cr.DownloadExcel(contractor, startVal, endVal), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
            }

            // Read the file into a byte array
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(fileBytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
        }
        public IActionResult DownloadReport(string reportType)
        {
            // Get the file path from the configuration
            string filePath = "";
            string fileName = "";
            string webRootPath = _webHostEnvironment.WebRootPath;


            switch (reportType)
            {
                case "HeliasDailyReport":
                    DailyReports dr = new DailyReports(_configuration);
                    //filePath = _configuration["FilePaths:HeliasDailyReport"];
                    fileName = "Helias_Sales_Report";
                    return File(dr.DownloadExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");

                case "Territory_Report":
                    {
                        var tr = new TerritoryReports(_configuration);
                        //filePath = _configuration["FilePaths:HeliasDailyReport"];
                        fileName = "Helias Sales Report - Territory";
                        return File(tr.DownloadExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "Territory_International":
                    {
                        var ir = new InternationalReports(_configuration);
                        //filePath = _configuration["FilePaths:HeliasDailyReport"];
                        fileName = "Helias Sales Report - International";
                        return File(ir.DownloadExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Detail_Oya":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Details OYA";
                        return File(ar.DownloadExcel("OYA", "Details"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Detail_Pino":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Details PINO";
                        return File(ar.DownloadExcel("PINO", "Details"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Summary_Oya":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary OYA";
                        return File(ar.DownloadExcel("OYA"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                case "AR_Aging_Summary_Pino":
                    {
                        AgingReport ar = new AgingReport(_configuration);
                        fileName = "Aging Summary PINO";
                        return File(ar.DownloadExcel("PINO"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
                default:
                    {
                        var contractor = reportType.Split('_')[1];
                        CommissionReport cr = new CommissionReport(_configuration);
                        fileName = "Commission Report - " + contractor;
                        return File(cr.DownloadExcel(contractor), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
                    }
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true), AllowAnonymous]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}