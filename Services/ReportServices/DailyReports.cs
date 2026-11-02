using GemBox.Spreadsheet;
using Intuit.Ipp.Data;
using Microsoft.Office.Interop.Excel;

//using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Org.BouncyCastle.Ocsp;
using OYA_Web_Application.Models.Data.Reports;
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Diagnostics.Contracts;
using System.IO;
//using System.Runtime.InteropServices;

namespace OYA_Web_Application.Services.ReportServices
{
    public class DailyReports
    {
        #region Declarations
        private readonly IConfiguration _configuration;
        static string path = "";
        static string outputPath = "";
        string connectionString = "";

        static string qodbcConnString = "DSN=Quickbooks Data;OLE DB Services=-2;";
        static DateTime now = DateTime.Now;
        static string logPath = @"C:\LOG\Helias Sales Report\" + DateTime.Now.ToString().Replace(":", "").Replace("/", "");
        //static string path = _configuration["TemplatePath"]; //@"E:\Jaffar\OYA_Project\Project\OYA_Web\Reports\HeliasReports\Templates\";
        static string template = "Helias Sales Report - Template v24 Contractor.xlsx";     // new
        static string mainSheetName = "OYASales2024Letter net sales";
        static string mainStartCell = "G2";
        static string accountsCol = "F";
        static string prevYearCol = "Z";
        static string year2DatePercentCol = "Y";
        static int internationalRow = 41;

        //static int salonsDirectRow = 38;    
        //static int publicDirectRow = salonsDirectRow + 2;   
        //static int allOyaRow = salonsDirectRow + 4;
        //static int pinoRow = salonsDirectRow + 5;
        //static int heliasRow = salonsDirectRow + 6;

        static int otherRow = 26;
        //static int usaRow = 35;
        static int canadaRow = 28;
        static int mexicoRow = 29;
        //static int internationalRow = 23;
        static int trinidadTobagoRow = 30;

        // Calculate the date range for the previous year
        DateTime startOfLastYear = new DateTime(DateTime.Now.Year - 1, 1, 1);
        DateTime endOfLastYear = new DateTime(DateTime.Now.Year - 1, DateTime.Now.Month, DateTime.Now.Day);

        // international
        static string internationalSheetName = "OYASales2024Letter net sales 2";
        static string internationalStartCell = "G2";

        static string canadaSheetName = "OYASls24 Canada table";
        static string westSheetName = "OYASls24 West table";
        static string centralSheetName = "OYASls24 Central table";
        static string eastSheetName = "OYASls24 East table";

        static string dateCell = "P1";
        static string startCell = "D3";

        static string canadaEndCell = "U47";
        static string westEndCell = "U41";
        static string centralEndCell = "U61";
        static string eastEndCell = "U65";

        static string[] oyaClosedList = new string[] {
                                        "OYA Affinity Beauty Concepts", "OYA Vintage Salon Supplies", "OYA Beauty Sales", "OYA GF Distributors", "OYA Independent SS", "OYA Kevin Ray Distributor Group",
                                        "OYA Lunica Beauty LLC", "OYA New England Salon Solutions", "OYA PA Beauty Supply",   "OYA Princess Beauty Supply","OYA Sweis",
                                        "OYA Hair Whisper", "OYA Vintage Canada", "OYA Beauty Van-Mexico"
        };

        static string[] oyaJobList = new string[] { 
                // OYA ISCS
                "OYA ISCS:ISCS Derek Moser", 
                //"OYA ISCS:ISCS Karmen Hadley", 
                "OYA ISCS:ISCS Michelle Fanning",
                // OYA M Beauty
                "OYA M Beauty:Andrea's",
                "OYA M Beauty:Ariel Derrick (Residence)",
                "OYA M Beauty:Breanna Taylor (Residence)",
                "OYA M Beauty:Elite Salon",
                "OYA M Beauty:Flaunt Salon",
                "OYA M Beauty:Fran Medina (Residence)",
                "OYA M Beauty:GAL Studios", 
               // "OYA M Beauty:Got Roots Salon",  
                "OYA M Beauty:Irene Fernandez",
                "OYA M Beauty:Jagged Edge",
                "OYA M Beauty:Karen Poulsen",
                "OYA M Beauty:Koolfx Salon at Mesa Royale",
                "OYA M Beauty:Linda Reese (Residence)",
                "OYA M Beauty:Marc Reid Salon @ Steven Paul Salon",
                "OYA M Beauty:Mira Bella Hair Salon",
                //"OYA M Beauty:OYA M Beauty - Sample Sub Account", 
                "OYA M Beauty:R Salon Day Spa - Donna Raffanti",
                "OYA M Beauty:Salon Boutique - Arrowhead", 
                    // "OYA M Beauty:Sassy Chic",             
                "OYA M Beauty:Schwahn Salon",
                "OYA M Beauty:Shaggy's",
                "OYA M Beauty:Shear Creations",
                "OYA M Beauty:Stasyss Salon",
                "OYA M Beauty:Sublime Salon",
                // OYA Unique Salon Concepts
                "OYA Unique Salon Concepts:Lisa Gagne",


                "OYA Uberbeauty808",
                "OYA C.K. ColorHouse",
                "OYA Store Professional",
                "OYA Store Consumer",
                "OYA M Beauty",
                "OYA Josephs College",
                "OYA Blaze Salon",
                "OYA ProSalon Distribution Group",
                "OYA Central States Supply",
                "OYA Elite Salon Distributors USA OHIO",
                "OYA Discover Salon Services",
                "OYA Mid America Beauty Supply",
                "OYA Economy",
                "OYA Elite Salon Services",
                "OYA Valley Beauty Supply",
                "OYA Unique Salon Concepts",
                "OYA Style Care",
                "OYA ISCS",
                "OYA Prime Salon Services",
                "OYA Aidan Laura Aine Ltd - ALA",
                "OYA Desbo Distribution Inc.",
                "OYA SOID Distributors, P.R.",
                "OYA De Bleu Beauty Supply",
                "OYA Binani Desbo Group",
                "OYA Braliza LLC",
                "OYA NCN Beauty LLC",
                "OYA NCN Beauty LLC Venezuela",
                "OYA NCN Beauty LLC Ecuador",
                "OYA Brothers Beauty Products"
        };

        //Application oXL = null;
        //ExcelWorkbook oWB = null;
        ExcelWorksheet oSheet = null;
        ExcelWorksheet sheetInternational = null;    // new
        ExcelWorksheet sheetCanada = null;
        ExcelWorksheet sheetWest = null;
        ExcelWorksheet sheetCentral = null;
        ExcelWorksheet sheetEast = null;


        static string[] oyaList = new string[] {
                "OYA Vintage Salon Supplies",
                "OYA Kevin Ray Distributor Group", //"OYA La Bella", 
                "OYA Uberbeauty808",//"OYA Salon Suppliers Inc." changed name to "OYA Uberbeauty808" on 2019-05-29;
                "OYA Josephs College",
                "OYA Blaze Salon",
                "OYA ProSalon Distribution Group",
                "OYA Central States Supply",
                "OYA Blast of Beauty",
                "OYA New England Salon Solutions",
                "OYA Elite Salon Services", "OYA Economy",
                "OYA ISCS", "OYA ISCS - Other", "ISCS Derek Moser", "ISCS Michelle Fanning",
                "OYA Prime Salon Services",
                "OYA Affinity Beauty Concepts",
                "OYA Hair Whisper", "OYA Vintage Canada",
                "OYA Beauty Van-Mexico",
                "OYA Independent SS", "OYA Sweis", "OYA PA Beauty Supply",
                "OYA GF Distributors", "OYA Princess Beauty Supply",
                "OYA Kmac2", "OYA Lunica Beauty LLC",
                "OYA Unique Salon Concepts",
                "OYA Mid America Beauty Supply",
                //"DS Salon Voche",
                "OYA Beauty Sales","OYA Style Care",
                "OYA  Aidan Laura Aine Ltd",
                "Total OYA M Beauty",
                "OYA M Beauty - Sample Sub Account",
                "Flaunt Salon",
                "Irene Fernandez",
                "Karen Poulsen",
                "Sassy Chic",
                "Schwahn Salon","OYA Jagged Edge"
        };


        List<string> oyaContractorList = new List<string> { };

        List<Distributor> distributorList = new List<Distributor>();
        List<Contractor> contractorList = new List<Contractor>();

        // do not modify=================================

        //===============================================

        List<Sales> contractorsThisMonthList = new List<Sales>();
        List<Sales> OverallOyaSales = new List<Sales>();
        List<Sales> OverallPinoSales = new List<Sales>();
        decimal internetThisMonthToDate = 0;
        decimal contractorsThisMonthToDate = 0;

        //List<Sales> oyaLastYearToDateList = new List<Sales>();
        //List<Sales> pinoLastYearToDateList = new List<Sales>();
        //List<Sales> otherLastYearToDateList = new List<Sales>();

        //decimal oyaLastYearToDate = 0;

        //decimal pinoLastYearToDate = 0;
        //decimal otherLastYearToDate = 0;

        int pinoCustomerCount = 0;
        byte[] content = null;

        #endregion

        public DailyReports(IConfiguration configuration)
        {
            _configuration = configuration;
            path = _configuration["TemplatePath"];
            outputPath = _configuration["ReportOutputPath"];
            connectionString = _configuration.GetConnectionString("DefaultConnection");
        }
        private void FetchData()
        {
            #region DB to Reports
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand("dbo.USP_GetSalesDataForReports", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();

                    // Add parameters
                    command.Parameters.Add("@thisYear", SqlDbType.Int).Value = DateTime.Now.Year;
                    command.Parameters.Add("@thisMonth", SqlDbType.Int).Value = DateTime.Now.Month;

                    SqlDataAdapter adapt = new SqlDataAdapter(command);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);

                    System.Data.DataTable overallSalesTable = ds.Tables[0];
                    foreach (DataRow _row in overallSalesTable.Rows)
                    {
                        Sales item = new Sales();

                        item.Customer = Convert.ToString(_row["Customer"]);
                        item.Class = Convert.ToString(_row["ClassRefFullName"]);
                        item.Month = Convert.ToInt32(_row["Month"]);
                        item.Year = Convert.ToInt32(_row["Year"]);
                        item.Rep = Convert.ToString(_row["SalesRepRefFullName"]);
                        item.CustomerType = Convert.ToString(_row["CustomerTypeRefFullName"]);
                        item.TxnDate = Convert.ToDateTime(_row["TxnDate"]);
                        if (_row["Amount"] != DBNull.Value)
                        {
                            item.Amount = Convert.ToDecimal(_row["Amount"]);
                        }
                        OverallOyaSales.Add(item);
                    }

                    overallSalesTable = ds.Tables[1];
                    foreach (DataRow _row in overallSalesTable.Rows)
                    {
                        Sales item = new Sales();

                        item.Customer = Convert.ToString(_row["Customer"]);
                        item.Class = Convert.ToString(_row["ClassRefFullName"]);
                        item.Month = Convert.ToInt32(_row["Month"]);
                        item.Year = Convert.ToInt32(_row["Year"]);
                        item.Rep = Convert.ToString(_row["SalesRepRefFullName"]);
                        item.CustomerType = Convert.ToString(_row["CustomerTypeRefFullName"]);
                        item.TxnDate = Convert.ToDateTime(_row["TxnDate"]);
                        if (_row["Amount"] != DBNull.Value)
                        {
                            item.Amount = Convert.ToDecimal(_row["Amount"]);
                        }
                        OverallPinoSales.Add(item);
                    }
                }
            }
            #endregion
        }

        private int GetRowByLabel(ExcelWorksheet sheet, string label)
        {
            for (int i = 1; i <= sheet.Rows.Count; i++)
            {
                var val = sheet.Cells["B" + i].Value?.ToString();
                if (!string.IsNullOrEmpty(val) &&
                    val.Equals(label, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            throw new Exception($"{label} row not found");
        }
        private static readonly Dictionary<int, (string col, string colTerritory)> MonthColumnMap =
            new Dictionary<int, (string, string)>
        {
            {1, ("L","G")}, {2, ("M","H")}, {3, ("N","I")}, {4, ("O","J")},
            {5, ("P","K")}, {6, ("Q","L")}, {7, ("R","M")}, {8, ("S","N")},
            {9, ("T","O")}, {10, ("U","P")}, {11, ("V","Q")}, {12, ("W","R")}
        };

        private void GenerateExcelv2()
        {
            var oWB = ExcelFile.Load(path + template);

            oSheet = oWB.Worksheets[0];
            sheetInternational = oWB.Worksheets[1];

            int salonsDirectRow = GetRowByLabel(oSheet, "Salons Direct");

            int publicDirectRow = salonsDirectRow + 2;
            int allOyaRow = salonsDirectRow + 4;
            int pinoRow = salonsDirectRow + 5;
            int heliasRow = salonsDirectRow + 6;

            string mainEndCell = prevYearCol + heliasRow;
            string internationalEndCell = prevYearCol + internationalRow;

            string allOyaAcctsCell = accountsCol + allOyaRow;
            string pinoAcctsCell = accountsCol + pinoRow;
            string heliasAcctsCell = accountsCol + heliasRow;

            distributorList = GetDistributor();
            contractorList = GetContractors();
            oyaContractorList = contractorList.Select(c => c.Code).ToList();

            var today = DateTime.Today;
            var startOfCurrentYear = new DateTime(today.Year, 1, 1);
            var startOfLastYear = new DateTime(today.Year - 1, 1, 1);
            var endOfLastYear = new DateTime(today.Year - 1, today.Month, today.Day);

            try
            {
                for (int month = 1; month <= now.Month; month++)
                {
                    if (!MonthColumnMap.ContainsKey(month)) continue;

                    var (col, colTerritory) = MonthColumnMap[month];

                    var currentMonthOya = OverallOyaSales
                        .Where(o => o.Month == month && o.Year == now.Year)
                        .ToList();

                    var currentMonthPino = OverallPinoSales
                        .Where(o => o.Month == month && o.Year == now.Year)
                        .ToList();

                    distributorList.ForEach(d =>
                    {
                        var targetSheet = d.Sheet == "Main" ? oSheet : sheetInternational;

                        var current = GetSalesByCustomer(d.Customer, currentMonthOya);
                        var previous = GetSalesByCustomer(d.Customer,
                            OverallOyaSales.Where(x => x.Year != now.Year).ToList());

                        var currentYTD = OverallOyaSales
                            .Where(s => s.Customer == d.Customer &&
                                        s.TxnDate >= startOfCurrentYear &&
                                        s.TxnDate <= today)
                            .Sum(s => s.Amount) ?? 0;

                        var previousYTD = OverallOyaSales
                            .Where(s => s.Customer == d.Customer &&
                                        s.TxnDate >= startOfLastYear &&
                                        s.TxnDate <= endOfLastYear)
                            .Sum(s => s.Amount) ?? 0;

                        decimal percentage = previousYTD == 0
                            ? 0
                            : ((currentYTD - previousYTD) / previousYTD) * 100;

                        targetSheet.Cells[col + d.Row].Value = current;
                        targetSheet.Cells["K" + d.Row].Value = previous;
                        targetSheet.Cells[year2DatePercentCol + d.Row].Value = Math.Round(percentage) + " %";
                        targetSheet.Cells[prevYearCol + d.Row].Value = previousYTD;
                    });

                    contractorList.ForEach(c =>
                    {
                        var current = OverallOyaSales
                            .Where(x => x.Rep == c.Code && x.Month == month && x.Year == now.Year)
                            .Sum(s => s.Amount) ?? 0;

                        var previous = OverallOyaSales
                            .Where(x => x.Rep == c.Code && x.Year != now.Year)
                            .Sum(s => s.Amount) ?? 0;

                        var prevYTD = OverallOyaSales
                            .Where(s => s.Rep == c.Code &&
                                        s.Year != now.Year &&
                                        s.TxnDate <= endOfLastYear)
                            .Sum(s => s.Amount) ?? 0;

                        sheetInternational.Cells[col + c.Row].Value = current;
                        sheetInternational.Cells["K" + c.Row].Value = previous;
                        sheetInternational.Cells[prevYearCol + c.Row].Value = prevYTD;
                    });

                    oSheet.Cells[col + pinoRow].Value = currentMonthPino.Sum(s => s.Amount);
                    oSheet.Cells[col + allOyaRow].Value = currentMonthOya.Sum(s => s.Amount);

                    oSheet.Cells[col + salonsDirectRow].Value = currentMonthOya
                        .Where(w => w.Class?.ToLower() == "salons direct")
                        .Sum(s => s.Amount);

                    oSheet.Cells[col + publicDirectRow].Value = currentMonthOya
                        .Where(w => (w.Class?.ToLower() == "public direct" || w.Class?.ToLower() == "amazon"))
                        .Sum(s => s.Amount);

                    pinoCustomerCount = currentMonthPino
                        .Select(x => x.Customer)
                        .Distinct()
                        .Count();

                    oSheet.Cells[allOyaAcctsCell].Value = distributorList.Count + " accts";
                    oSheet.Cells[pinoAcctsCell].Value = pinoCustomerCount + " accts";
                    oSheet.Cells[heliasAcctsCell].Value = (distributorList.Count + pinoCustomerCount) + " accts";
                }

                var lastYear = OverallOyaSales.Where(o => o.Year != now.Year);

                oSheet.Cells[prevYearCol + allOyaRow].Value = lastYear
                    .Where(o => o.TxnDate <= endOfLastYear)
                    .Sum(s => s.Amount);

                oSheet.Cells[prevYearCol + pinoRow].Value = OverallPinoSales
                    .Where(o => o.Year != now.Year && o.TxnDate <= endOfLastYear)
                    .Sum(s => s.Amount);

                oSheet.Cells["K" + allOyaRow].Value = lastYear.Sum(s => s.Amount);

                using (var ms = new MemoryStream())
                {
                    oWB.Save(ms, SaveOptions.XlsxDefault);
                    content = ms.ToArray();
                }
            }
            catch
            {
                // You were swallowing exception already — kept same behavior
            }
        }
        private void GenerateExcel()
        {
            int thisMonth = now.Month;
            //string col = "";
            //string colTerritory;

            var oWB = ExcelFile.Load(path + template);

            oSheet = oWB.Worksheets[0];
            sheetInternational = oWB.Worksheets[1];

            int salonsDirectRow = GetRowByLabel(oSheet, "Salons Direct");

            int publicDirectRow = salonsDirectRow + 2;
            int allOyaRow = salonsDirectRow + 4;
            int pinoRow = salonsDirectRow + 5;
            int heliasRow = salonsDirectRow + 6;

            string mainEndCell = prevYearCol + heliasRow;
            string internationalEndCell = prevYearCol + internationalRow;

            string allOyaAcctsCell = accountsCol + allOyaRow;
            string pinoAcctsCell = accountsCol + pinoRow;
            string heliasAcctsCell = accountsCol + heliasRow;

            distributorList = GetDistributor().Where(d => d.Row > 0).OrderBy(d => d.Row).ToList();
            contractorList = GetContractors();
            oyaContractorList = contractorList.Select(c => c.Code).ToList();

            try
            {
                for (int i = 1; i <= now.Month; i++)
                {
                    thisMonth = i;

                    //CompareCustomerTotals(OverallOyaSales, thisMonth, now.Year);
                    //FindExactMissing309(OverallOyaSales, thisMonth, now.Year);
                    //AuditTransactionClassification(OverallOyaSales, thisMonth, now.Year);
                    //SummarizeClassificationIssues(OverallOyaSales, thisMonth, now.Year);

                    var (col, colTerritory) = MonthColumnMap[thisMonth];
                    var today = DateTime.Today;
                    DateTime startOfCurrentYear = new DateTime(today.Year, 1, 1);
                    DateTime startOfLastYear = new DateTime(today.Year - 1, 1, 1);
                    DateTime endOfLastYear = new DateTime(today.Year - 1, today.Month, today.Day);

                    var currentMonthData = OverallOyaSales
                        .Where(o => o.Month == thisMonth && o.Year == now.Year)
                        .ToList();

                    var lastYearData = OverallOyaSales
                        .Where(o => o.Year != now.Year)
                        .ToList();

                    var currentYtdData = OverallOyaSales
                        .Where(s => s.TxnDate >= startOfCurrentYear && s.TxnDate <= today)
                        .ToList();

                    var previousYtdData = OverallOyaSales
                        .Where(s => s.TxnDate >= startOfLastYear && s.TxnDate <= endOfLastYear)
                        .ToList();

                    distributorList.ForEach(d =>
                    {
                        var sheet = d.Sheet == "Main" ? oSheet : sheetInternational;

                        sheet.Cells[col + d.Row].Value =
                            GetSalesByCustomer(d.Customer,
                                currentMonthData.Where(o => o.Customer.ToLower() == d.Customer.ToLower()).ToList());

                        sheet.Cells["K" + d.Row].Value =
                            GetSalesByCustomer(d.Customer,
                                lastYearData.Where(o => o.Customer.ToLower() == d.Customer.ToLower()).ToList());
                        var currentYearToDateSales = currentYtdData
                            .Where(s => s.Customer == d.Customer)
                            .Sum(s => (decimal?)s.Amount) ?? 0;

                        var previousYearToDateSales = previousYtdData
                            .Where(s => s.Customer == d.Customer)
                            .Sum(s => (decimal?)s.Amount) ?? 0;

                        decimal percentageDifference = 0;

                        if (previousYearToDateSales != 0)
                        {
                            percentageDifference =
                                ((currentYearToDateSales - previousYearToDateSales) / previousYearToDateSales) * 100;
                        }

                        sheet.Cells[year2DatePercentCol + d.Row].Value =
                            Math.Round(percentageDifference) + " %";

                        sheet.Cells[prevYearCol + d.Row].Value = previousYearToDateSales;
                    });

                    contractorList.ForEach(c =>
                    {
                        var currentValue = currentMonthData
                            .Where(x => x.Rep.Equals(c.Code))
                            .Sum(s => (decimal?)s.Amount) ?? 0;

                        sheetInternational.Cells[col + c.Row].Value = currentValue;

                        var previousYearToDateSales = previousYtdData
                            .Where(s => s.Rep.Equals(c.Code))
                            .Sum(s => (decimal?)s.Amount) ?? 0;

                        sheetInternational.Cells[prevYearCol + c.Row].Value = previousYearToDateSales;

                        var lastYearTotal = lastYearData
                            .Where(x => x.Rep.Equals(c.Code))
                            .Sum(s => (decimal?)s.Amount) ?? 0;

                        sheetInternational.Cells["K" + c.Row].Value = lastYearTotal;
                    });

                    oSheet.Cells[col + pinoRow].Value = OverallPinoSales.Where(w => w.Month == thisMonth
                    && w.Year == now.Year).Sum(s => s.Amount);

                    oSheet.Cells[col + allOyaRow].Value = OverallOyaSales.Where(w => w.Month == thisMonth
                    && w.Year == now.Year).Sum(s => s.Amount);

                    var salonsDirectSales = OverallOyaSales
                        .Where(w => w.Class.ToLower() == "salons direct"
                                 && w.Month == thisMonth
                                 && w.Year == now.Year)
                        .Sum(s => (decimal?)s.Amount) ?? 0;

                    var publicDirectSales = OverallOyaSales
                        .Where(w => w.Class.ToLower() == "public direct"
                                 && w.Month == thisMonth
                                 && w.Year == now.Year)
                        .Sum(s => (decimal?)s.Amount) ?? 0;

                    var amazonSales = OverallOyaSales
                        .Where(w => w.Class.ToLower() == "amazon"
                                 && w.Month == thisMonth
                                 && w.Year == now.Year)
                        .Sum(s => (decimal?)s.Amount) ?? 0;

                    oSheet.Cells[col + salonsDirectRow].Value = salonsDirectSales;

                    oSheet.Cells[col + publicDirectRow].Value = publicDirectSales + amazonSales;

                    oSheet.Cells[col + allOyaRow].Value = OverallOyaSales.Where(o => o.Month == thisMonth && o.Year == now.Year).Sum(s => s.Amount);

                    oSheet.Cells[col + pinoRow].Value = OverallPinoSales.Where(o => o.Month == thisMonth && o.Year == now.Year).Sum(s => s.Amount);

                    pinoCustomerCount = OverallPinoSales.DistinctBy(x => x.Customer).Count();

                    oSheet.Cells[allOyaAcctsCell].Value = distributorList.Count() + " accts";
                    oSheet.Cells[pinoAcctsCell].Value = pinoCustomerCount + "  accts";
                    oSheet.Cells[heliasAcctsCell].Value = distributorList.Count() + pinoCustomerCount + " accts";

                }


                #region Prev Year to Date

                var lastYearOya = OverallOyaSales
                    .Where(o => o.Year != now.Year && o.TxnDate <= endOfLastYear)
                    .ToList();

                var lastYearPino = OverallPinoSales
                    .Where(o => o.Year != now.Year && o.TxnDate <= endOfLastYear)
                    .ToList();

                oSheet.Cells[prevYearCol + allOyaRow].Value =
                    lastYearOya.Sum(s => (decimal?)s.Amount) ?? 0;

                oSheet.Cells[prevYearCol + pinoRow].Value =
                    lastYearPino.Sum(s => (decimal?)s.Amount) ?? 0;

                var lastYearClassWise = lastYearOya
                    .GroupBy(w => w.Class == null ? "" : w.Class.ToLower())
                    .ToDictionary(g => g.Key, g => g.Sum(x => (decimal?)x.Amount) ?? 0);

                var salonsDirectLastYear = lastYearClassWise.GetValueOrDefault("salons direct");
                var publicDirectLastYear = lastYearClassWise.GetValueOrDefault("public direct");
                var amazonLastYear = lastYearClassWise.GetValueOrDefault("amazon");

                oSheet.Cells[prevYearCol + salonsDirectRow].Value = salonsDirectLastYear;
                oSheet.Cells[prevYearCol + publicDirectRow].Value = publicDirectLastYear + amazonLastYear;

                #endregion

                oSheet.Cells["K" + salonsDirectRow].Value = salonsDirectLastYear;
                oSheet.Cells["K" + publicDirectRow].Value = publicDirectLastYear + amazonLastYear;

                oSheet.Cells["K" + pinoRow].Value =
                    lastYearPino.Sum(s => (decimal?)s.Amount) ?? 0;

                oSheet.Cells["K" + allOyaRow].Value =
                    lastYearOya.Sum(s => (decimal?)s.Amount) ?? 0;

                if (now.Year == 2026)
                {
                    #region Prev Year to Date

                    oSheet.Cells[prevYearCol + allOyaRow].Value = OverallOyaSales.Where(o => o.TxnDate <= endOfLastYear && o.Year != now.Year).Sum(s => s.Amount);
                    oSheet.Cells[prevYearCol + pinoRow].Value = OverallPinoSales.Where(o => o.TxnDate <= endOfLastYear && o.Year != now.Year).Sum(s => s.Amount);

                    oSheet.Cells[prevYearCol + publicDirectRow].Value = OverallOyaSales.Where(w => w.Class.ToLower() == "internet order" && w.CustomerType.ToLower() != "other"
                                                                            && w.Year != now.Year && w.TxnDate <= endOfLastYear).Sum(s => s.Amount);

                    oSheet.Cells[prevYearCol + salonsDirectRow].Value = OverallOyaSales.Where(w => w.CustomerType.ToLower() == "other"
                                                                     && w.Year != now.Year && w.TxnDate <= endOfLastYear).Sum(s => s.Amount);


                    #endregion

                    var lastYear = OverallOyaSales.Where(o => o.Year != now.Year);

                    oSheet.Cells["K" + salonsDirectRow].Value = lastYear.Where(w => w.CustomerType.ToLower() == "other").Sum(s => s.Amount);
                    oSheet.Cells["K" + publicDirectRow].Value = lastYear.Where(w => w.Class.ToLower() == "internet order" && w.CustomerType.ToLower() != "other").Sum(s => s.Amount);
                    oSheet.Cells["K" + pinoRow].Value = OverallPinoSales.Where(w => w.Year != now.Year).Sum(s => s.Amount);
                    oSheet.Cells["K" + allOyaRow].Value = lastYear.Sum(s => s.Amount);
                }

                using (var memoryStream = new MemoryStream())
                {
                    oWB.Save(memoryStream, SaveOptions.XlsxDefault);

                    content = memoryStream.ToArray();
                }

            }
            catch (Exception ex)
            {
                //throw;
            }
            finally
            {

            }
        }
        public byte[] DownloadExcel()
        {
            FetchData();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            GenerateExcel();


            return content;
        }

        static decimal GetSalesByCustomer(string customer, IList<Sales> list)
        {
            decimal value = 0;
            var filteredList = list.Where(x => x.Customer.ToLower().Equals(customer.ToLower())).ToList();
            if (filteredList.Count() > 0)
                value = filteredList.Sum(s => s.Amount) ?? 0;

            return value;
        }
        public List<Contractor> GetContractors()
        {
            List<Contractor> contractorList = new List<Contractor>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand com = new SqlCommand("SELECT Code, Row, Sheet, Contractor FROM Contractor", conn))
                {
                    SqlDataReader reader = com.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader["Code"] != DBNull.Value)
                        {
                            contractorList.Add(new Contractor
                            {
                                Code = reader["Code"].ToString(),
                                Row = Convert.ToInt32(reader["Row"]),
                                Sheet = reader["Sheet"].ToString(),
                                Label = reader["Contractor"].ToString(),
                            });
                        }
                    }
                    reader.Close(); // Close the SqlDataReader after reading data
                }
                conn.Close(); // Close the SqlConnection after usage
            }
            return contractorList;
        }

        public List<Distributor> GetDistributor()
        {
            var distributorList = new List<Distributor>();

            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (var com = new SqlCommand("SELECT Customer, Sheet, Row FROM Distributor WHERE IsActive = 1 ORDER BY Sheet, Row", conn))
                {
                    var reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        distributorList.Add(new Distributor
                        {
                            Customer = Convert.ToString(reader["Customer"]),
                            Row = Convert.ToInt32(reader["Row"]),
                            Sheet = reader["Sheet"].ToString(),
                        });
                    }
                    reader.Close();
                }
                conn.Close();
            }
            return distributorList;
        }
        private static string Normalize(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            var value = input.ToLower().Trim();

            // remove everything after comma
            var commaIndex = value.IndexOf(",");
            if (commaIndex > 0)
                value = value.Substring(0, commaIndex);

            // remove extra spaces
            value = System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ");

            return value.Trim();
        }
        private static readonly Dictionary<string, decimal> QBSalonsData =
     new Dictionary<string, decimal>
 {
    { Normalize("Alexandra Shirzad Shirzad"), 52.00m },
    { Normalize("Ashley Kowalczyk"), 303.75m },
    { Normalize("Beach Bum Tanz LLC MILLER"), 25.00m },
    { Normalize("Becky Box"), 226.15m },
    { Normalize("Brandie Rounds"), 25.00m },
    { Normalize("Chem-Steel Distro Roger Cribb"), 162.44m },
    { Normalize("CHESNEY TIPTON Tipton"), 175.00m },
    { Normalize("Colton Gieselman"), 35.85m },
    { Normalize("Deb parker Deb parker"), 9.50m },
    { Normalize("Denee Clifford"), 71.70m },
    { Normalize("Diana S Kurzka"), 21.00m },
    { Normalize("DS Embodyment Salon & Spa"), 248.50m },
    { Normalize("DS Hair Venue"), 146.00m },
    { Normalize("DS Jordan Deisbeck"), 455.50m },
    { Normalize("DS Spa Isbell"), 114.35m },
    { Normalize("DSCA InnerChi, LLC"), 790.65m },
    { Normalize("DSCA Joni Nauton"), 183.40m },
    { Normalize("DSCA Penny Colbert"), 119.50m },
    { Normalize("DSCO Tanya Auville"), 269.05m },
    { Normalize("DSGA Shawon Pinder"), 288.00m },
    { Normalize("DSID Rebecca Williams"), 334.30m },
    { Normalize("DSNV Chris The Hair Guy"), 463.74m },
    { Normalize("DSUT Steur and Friends"), 166.40m },
    { Normalize("Jillian Vissa"), 76.45m },
    { Normalize("Julie Brahier"), 420.65m },
    { Normalize("karen axsom"), 69.00m },
    { Normalize("Karen Baez Baez"), 209.05m },
    { Normalize("Katie Barton Katie Barton"), 116.00m },
    { Normalize("Marina Guerrero"), 372.00m },
    { Normalize("marladietch@hotmail.com"), 57.30m },
    { Normalize("OLIVIA HOWELL-CORA & COMPANY OLIVIA HOWEL"), 57.00m },
    { Normalize("Patricia Wheat"), 15.50m },
    { Normalize("Ricardo Moya"), 0.00m },
    { Normalize("Tiffany Sears"), 62.50m },
    { Normalize("TLP HOLDINGS LLC POULIOT"), 30.00m },
    { Normalize("Trade Show - OYA"), 10322.00m },
    { Normalize("TRANSITIONS OF INDIANA-INDIANAPOLIS Brack"), 198.00m },
    { Normalize("Trudy Peterson"), 252.65m },
    { Normalize("Virginia Cady"), 138.00m },
    { Normalize("Yeny Marcela Garcia"), 478.00m }
 };
        public static void CompareCustomerTotals(
            List<Sales> overallOyaSales,
            int thisMonth,
            int year)
        {
            var systemData = overallOyaSales
                .Where(x => x.Month == thisMonth && x.Year == year)
                .ToList();

            var systemGrouped = systemData
                .GroupBy(x => Normalize(x.Customer))
                .ToDictionary(
                    g => g.Key,
                    g => new
                    {
                        Total = g.Sum(x => (decimal?)x.Amount) ?? 0,
                        Count = g.Count(),
                        MinDate = g.Min(x => x.TxnDate),
                        MaxDate = g.Max(x => x.TxnDate),
                        Classes = g.GroupBy(x => x.Class ?? "NULL")
                                   .ToDictionary(cg => cg.Key, cg => cg.Sum(x => x.Amount))
                    }
                );

            Console.WriteLine("Customer | QB | System | Diff | Count | DateRange");
            Console.WriteLine("------------------------------------------------------------------");

            decimal totalDiff = 0;

            foreach (var qb in QBSalonsData)
            {
                var exists = systemGrouped.ContainsKey(qb.Key);
                var systemValue = exists ? systemGrouped[qb.Key].Total : 0;
                var diff = qb.Value - systemValue;

                totalDiff += diff;

                if (diff != 0)
                {
                    var count = exists ? systemGrouped[qb.Key].Count : 0;
                    var minDate = exists ? systemGrouped[qb.Key].MinDate.ToShortDateString() : "-";
                    var maxDate = exists ? systemGrouped[qb.Key].MaxDate.ToShortDateString() : "-";

                    Console.WriteLine($"{qb.Key} | {qb.Value} | {systemValue} | {diff} | {count} | {minDate} - {maxDate}");

                    // class breakdown (this is gold for your case)
                    if (exists)
                    {
                        Console.WriteLine("   Class Breakdown:");
                        foreach (var cls in systemGrouped[qb.Key].Classes)
                        {
                            Console.WriteLine($"      {cls.Key} : {cls.Value}");
                        }
                    }
                }
            }

            Console.WriteLine("------------------------------------------------------------------");
            Console.WriteLine("TOTAL DIFF: " + totalDiff);
            Console.WriteLine("Month: " + thisMonth);

            // --- EXTRA: system customers NOT in QB ---
            var qbKeys = QBSalonsData.Keys.ToHashSet();

            var extraCustomers = systemGrouped
                .Where(x => !qbKeys.Contains(x.Key))
                .Select(x => new
                {
                    Customer = x.Key,
                    Total = x.Value.Total,
                    Count = x.Value.Count
                })
                .OrderByDescending(x => x.Total)
                .Take(20)
                .ToList();

            Console.WriteLine("\nTop System Customers NOT in QB:");
            foreach (var item in extraCustomers)
            {
                Console.WriteLine($"{item.Customer} | {item.Total} | {item.Count}");
            }

            // --- EXTRA: QB customers missing in system ---
            var missingInSystem = QBSalonsData
                .Where(qb => !systemGrouped.ContainsKey(qb.Key))
                .ToList();

            Console.WriteLine("\nQB Customers Missing in System:");
            foreach (var item in missingInSystem)
            {
                Console.WriteLine($"{item.Key} | {item.Value}");
            }
        }
        public static void FindExactMissing309(
    List<Sales> overallOyaSales,
    int thisMonth,
    int year)
        {
            var data = overallOyaSales
                .Where(x => x.Month == thisMonth && x.Year == year)
                .ToList();

            var matches = data
                .Where(x =>
                {
                    var normalized = Normalize(x.Customer);
                    var cls = x.Class?.Trim().ToLower();

                    return QBSalonsData.ContainsKey(normalized)
                           && string.IsNullOrWhiteSpace(cls);
                })
                .Select(x => new
                {
                    x.Customer,
                    x.Amount,
                    x.Class,
                    x.TxnDate
                })
                .ToList();

            var total = matches.Sum(x => (decimal?)x.Amount) ?? 0;

            Console.WriteLine("---- EXACT MISSING (SHOULD BE SALONS) ----");

            foreach (var m in matches)
            {
                Console.WriteLine($"{m.Customer} | {m.Amount} | {m.Class ?? "NULL"} | {m.TxnDate}");
            }

            Console.WriteLine("------------------------------------------");
            Console.WriteLine("TOTAL: " + total);
        }
        public static void AuditTransactionClassification(
    List<Sales> overallOyaSales,
    int thisMonth,
    int year)
        {
            var data = overallOyaSales
                .Where(x => x.Month == thisMonth && x.Year == year)
                .ToList();

            Console.WriteLine("---- POTENTIAL CLASSIFICATION ISSUES ----");
            Console.WriteLine("Customer | Amount | Class | Issue");
            Console.WriteLine("----------------------------------------------------------");

            foreach (var x in data)
            {
                var normalizedCustomer = Normalize(x.Customer);
                var cls = x.Class?.Trim().ToLower();

                string issue = null;

                // Case 1: No class
                if (string.IsNullOrWhiteSpace(cls))
                {
                    issue = "NULL CLASS";
                }
                // Case 2: Unknown class
                else if (cls != "salons direct" &&
                         cls != "public direct" &&
                         cls != "amazon")
                {
                    issue = "UNKNOWN CLASS";
                }

                // Case 3: Belongs to QB salons but class says otherwise
                if (QBSalonsData.ContainsKey(normalizedCustomer))
                {
                    if (cls != "salons direct")
                    {
                        issue = (issue == null ? "" : issue + " | ") + "SHOULD BE SALONS";
                    }
                }

                if (issue != null)
                {
                    Console.WriteLine($"{x.Customer} | {x.Amount} | {x.Class ?? "NULL"} | {issue}");
                }
            }
        }
        public static void SummarizeClassificationIssues(
    List<Sales> overallOyaSales,
    int thisMonth,
    int year)
        {
            var data = overallOyaSales
                .Where(x => x.Month == thisMonth && x.Year == year)
                .ToList();

            var issues = data
                .Select(x =>
                {
                    var normalizedCustomer = Normalize(x.Customer);
                    var cls = x.Class?.Trim().ToLower();

                    string issue = null;

                    if (string.IsNullOrWhiteSpace(cls))
                        issue = "NULL CLASS";
                    else if (cls != "salons direct" &&
                             cls != "public direct" &&
                             cls != "amazon")
                        issue = "UNKNOWN CLASS";

                    if (QBSalonsData.ContainsKey(normalizedCustomer) && cls != "salons direct")
                        issue = (issue == null ? "" : issue + " | ") + "SHOULD BE SALONS";

                    return new
                    {
                        x.Customer,
                        x.Amount,
                        Issue = issue
                    };
                })
                .Where(x => x.Issue != null)
                .GroupBy(x => x.Issue)
                .Select(g => new
                {
                    Issue = g.Key,
                    Total = g.Sum(x => x.Amount),
                    Count = g.Count()
                })
                .OrderByDescending(x => x.Total)
                .ToList();

            Console.WriteLine("\n---- ISSUE SUMMARY ----");
            foreach (var i in issues)
            {
                Console.WriteLine($"{i.Issue} | {i.Total} | {i.Count}");
            }
        }
    }
    public class Sales
    {
        public string Customer { get; set; }
        public Nullable<decimal> Amount { get; set; }
        public string RowType { get; set; }
        public string Class { get; set; }
        public string Rep { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public DateTime TxnDate { get; set; }
        public string CustomerType { get; set; }
        public string AccountRefFullName { get; set; }
    }

    public class SalesByCustomer
    {
        public string Customer { get; set; }
        public string Amount { get; set; }
        //public int Year { get; set; }
        //public int Month { get; set; }
        //public string ShipState { get; set; }
    }
}
