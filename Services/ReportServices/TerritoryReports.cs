using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using OYA_Web_Application.Models.Data.Reports;
using System.Data;
using System.Data.SqlClient;

namespace OYA_Web_Application.Services.ReportServices
{
    public class TerritoryReports
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
        static string template = "Helias Sales Report - Territory.xlsx";      // new
        static string mainSheetName = "OYASales2024Letter net sales";
        static string mainStartCell = "G2";
        static string accountsCol = "F";
        static string prevYearCol = "Z";
        static int internationalRow = 41;
        static int salonsDirectRow = 35;     // changed from Other
        static int publicDirectRow = 37;    // Internet Order including Am  azon
        static int allOyaRow = 39;
        static int pinoRow = 40;
        static int heliasRow = 41;
        static int otherRow = 26;
        //static int usaRow = 35;
        static int canadaRow = 28;
        static int mexicoRow = 29;
        //static int internationalRow = 23;
        static int trinidadTobagoRow = 30;

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
        string mainEndCell = prevYearCol + heliasRow;
        string internationalEndCell = prevYearCol + internationalRow;

        string allOyaAcctsCell = accountsCol + allOyaRow;
        string pinoAcctsCell = accountsCol + pinoRow;
        string heliasAcctsCell = accountsCol + heliasRow;
        //===============================================
        List<Sales> oyaThisMonthList = new List<Sales>();
        List<Sales> pinoThisMonthList = new List<Sales>();
        List<Sales> otherThisMonthList = new List<Sales>();
        List<Sales> contractorsThisMonthList = new List<Sales>();
        List<Sales> OverallSales = new List<Sales>();

        decimal oyaThisMonthToDate = 0;
        decimal pinoThisMonthToDate = 0;
        decimal otherThisMonthToDate = 0;
        decimal internetThisMonthToDate = 0;
        decimal contractorsThisMonthToDate = 0;

        List<Sales> oyaLastYearToDateList = new List<Sales>();
        List<Sales> pinoLastYearToDateList = new List<Sales>();
        List<Sales> otherLastYearToDateList = new List<Sales>();

        decimal oyaLastYearToDate = 0;
        //decimal oyaUSALastYearToDate = 0;
        //decimal oyaCanadaLastYearToDate = 0;

        decimal pinoLastYearToDate = 0;
        decimal otherLastYearToDate = 0;

        int pinoCustomerCount = 0;
        byte[] content = null;

        #endregion

        public TerritoryReports(IConfiguration configuration)
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
                        item.TxnDate = Convert.ToDateTime(_row["TxnDate"]);
                        item.Year = Convert.ToInt32(_row["Year"]);
                        item.Rep = Convert.ToString(_row["SalesRepRefFullName"]);
                        item.CustomerType = Convert.ToString(_row["CustomerTypeRefFullName"]);
                        if (_row["Amount"] != DBNull.Value)
                        {
                            item.Amount = Convert.ToDecimal(_row["Amount"]);
                        }
                        OverallSales.Add(item);
                    }
                }
            }
            #endregion
        }
        private void GenerateExcel()
        {
            int thisMonth = now.Month;
            string col = "";
            string colTerritory;

            var oWB = ExcelFile.Load(path + template);

            sheetWest = oWB.Worksheets[0];
            sheetCentral = oWB.Worksheets[1];
            sheetEast = oWB.Worksheets[2];

            distributorList = GetDistributor();
            //contractorList = GetContractors();
            //oyaContractorList = contractorList.Select(c => c.Code).ToList();

            var today = DateTime.Today;
            DateTime startOfCurrentYear = new DateTime(today.Year, 1, 1);
            DateTime startOfLastYear = new DateTime(today.Year - 1, 1, 1);
            DateTime endOfLastYear = new DateTime(today.Year - 1, today.Month, today.Day);

            sheetWest.Cells["A" + 1].Value = $"Net sales {DateTime.Now.Year} in US$";
            sheetWest.Cells["P" + 1].Value = $"{DateTime.Now.ToString("MMMM dd, yyyy")}";


            sheetCentral.Cells["A" + 1].Value = $"Net sales {DateTime.Now.Year} in US$";
            sheetCentral.Cells["P" + 1].Value = $"{DateTime.Now.ToString("MMMM dd, yyyy")}";


            sheetEast.Cells["A" + 1].Value = $"Net sales {DateTime.Now.Year} in US$";
            sheetEast.Cells["P" + 1].Value = $"{DateTime.Now.ToString("MMMM dd, yyyy")}";
            try
            {
                for (int i = 1; i <= 12; i++)
                {
                    thisMonth = i;
                    switch (thisMonth)
                    {
                        case 1:
                            col = "L";
                            colTerritory = "G";
                            break;
                        case 2:
                            col = "M";
                            colTerritory = "H";
                            break;
                        case 3:
                            col = "N";
                            colTerritory = "I";
                            break;
                        case 4:
                            col = "O";
                            colTerritory = "J";
                            break;
                        case 5:
                            col = "P";
                            colTerritory = "K";
                            break;
                        case 6:
                            col = "Q";
                            colTerritory = "L";
                            break;
                        case 7:
                            col = "R";
                            colTerritory = "M";
                            break;
                        case 8:
                            col = "S";
                            colTerritory = "N";
                            break;
                        case 9:
                            col = "T";
                            colTerritory = "O";
                            break;
                        case 10:
                            col = "U";
                            colTerritory = "P";
                            break;
                        case 11:
                            col = "V";
                            colTerritory = "Q";
                            break;
                        case 12:
                            col = "W";
                            colTerritory = "R";
                            break;
                        default:
                            colTerritory = "G";
                            break;
                    }
                    distributorList.Where(d => d.WestRow > 0).ToList().ForEach(d =>
                    {
                        if (thisMonth <= now.Month)
                        {
                            var currSales = OverallSales.Where(o => o.Month == thisMonth && o.Year == now.Year).ToList();
                            sheetWest.Cells[colTerritory + (d.WestRow + 1)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year == now.Year).ToList());

                            var currentYearToDateSales = OverallSales
                                    .Where(s => s.TxnDate >= startOfCurrentYear && s.TxnDate <= today && s.Customer.ToLower() == d.Customer.ToLower())
                                    .Sum(s => s.Amount);


                            var previousYearToDateSales = OverallSales
                                .Where(s => s.TxnDate >= startOfLastYear && s.TxnDate <= endOfLastYear && s.Customer.ToLower() == d.Customer.ToLower())
                                .Sum(s => s.Amount);

                            sheetWest.Cells["U" + d.WestRow].Value = previousYearToDateSales;

                            decimal percentageDifference = 0;

                            if (previousYearToDateSales != 0)
                            {
                                percentageDifference = (((decimal)currentYearToDateSales - (decimal)previousYearToDateSales) / (decimal)previousYearToDateSales) * 100;
                            }

                            sheetWest.Cells["T" + (d.WestRow + 1)].Value = Math.Round(percentageDifference) + " %";
                        }
                        sheetWest.Cells[colTerritory + (d.WestRow)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year != now.Year).ToList());
                    });
                    distributorList.Where(d => d.CentralRow > 0).ToList().ForEach(d =>
                    {
                        if (thisMonth <= now.Month)
                        {
                            sheetCentral.Cells[colTerritory + (d.CentralRow + 1)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year == now.Year).ToList());
                            var currentYearToDateSales = OverallSales
                                  .Where(s => s.TxnDate >= startOfCurrentYear && s.TxnDate <= today && s.Customer.ToLower() == d.Customer.ToLower())
                                  .Sum(s => s.Amount);


                            var previousYearToDateSales = OverallSales
                                .Where(s => s.TxnDate >= startOfLastYear && s.TxnDate <= endOfLastYear && s.Customer.ToLower() == d.Customer.ToLower())
                                .Sum(s => s.Amount);

                            sheetCentral.Cells["U" + d.CentralRow].Value = previousYearToDateSales;
                            decimal percentageDifference = 0;

                            if (previousYearToDateSales != 0)
                            {
                                percentageDifference = (((decimal)currentYearToDateSales - (decimal)previousYearToDateSales) / (decimal)previousYearToDateSales) * 100;
                            }

                            sheetCentral.Cells["T" + (d.CentralRow + 1)].Value = Math.Round(percentageDifference) + " %";
                        }
                        sheetCentral.Cells[colTerritory + (d.CentralRow)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year != now.Year).ToList());
                    });
                    distributorList.Where(d => d.EastRow > 0).ToList().ForEach(d =>
                    {
                        if (thisMonth <= now.Month)
                        {
                            sheetEast.Cells[colTerritory + (d.EastRow + 1)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year == now.Year).ToList());
                            var currentYearToDateSales = OverallSales
                                   .Where(s => s.TxnDate >= startOfCurrentYear && s.TxnDate <= today && s.Customer.ToLower() == d.Customer.ToLower())
                                   .Sum(s => s.Amount);
                            


                            var previousYearToDateSales = OverallSales
                                .Where(s => s.TxnDate >= startOfLastYear && s.TxnDate <= endOfLastYear && s.Customer.ToLower() == d.Customer.ToLower())
                                .Sum(s => s.Amount);

                            sheetEast.Cells["U" + d.EastRow].Value = previousYearToDateSales;

                            decimal percentageDifference = 0;

                            if (previousYearToDateSales != 0)
                            {
                                percentageDifference = (((decimal)currentYearToDateSales - (decimal)previousYearToDateSales) / (decimal)previousYearToDateSales) * 100;
                            }

                            sheetEast.Cells["T" + (d.EastRow + 1)].Value = Math.Round(percentageDifference) + " %";
                        }
                        sheetEast.Cells[colTerritory + (d.EastRow)].Value = GetSalesByCustomer(d.Customer, OverallSales.Where(o => o.Month == thisMonth && o.Year != now.Year).ToList());
                    });
                }
                sheetWest.Cells["P1"].Value = DateTime.Now.ToString("MMMM dd, yyyy");
                sheetCentral.Cells["P1"].Value = DateTime.Now.ToString("MMMM dd, yyyy");
                sheetEast.Cells["P1"].Value = DateTime.Now.ToString("MMMM dd, yyyy");
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
            SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.Stop;
            GenerateExcel();


            return content;
        }

        static void CopyToDailyReport(Application sourceXL, Workbook sourceWb, string sourceWorksheetName, string rep, string begCell, string endCell)
        {
            try
            {
                Object defaultArg = Type.Missing;
                Workbook copyWb = null;
                Worksheet copyWs = null;
                //string begCell = "E2";

                string fileName = "";
                string date = now.ToString("MMMM dd, yyyy");

                copyWb = sourceXL.Workbooks.Open(path + "Helias Sales Report" + fileName + ".xlsx",
                            0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //copyWs = (Worksheet)copyWb.ActiveSheet;
                copyWs = (Worksheet)copyWb.Worksheets[sourceWorksheetName];
                if (now.Day < 10)
                {
                    copyWs.PageSetup.RightHeader = "&" + 12 + "&\"" + "Arial" + "\"" + "per " + now.ToString("MMMM d, yyyy");
                }
                else
                {
                    copyWs.PageSetup.RightHeader = "&" + 12 + "&\"" + "Arial" + "\"" + "per " + now.ToString("MMMM d, yyyy");
                }

                ((Worksheet)sourceWb.Worksheets[sourceWorksheetName]).get_Range(begCell, endCell).Copy(Type.Missing);
                copyWs.get_Range(begCell, endCell).PasteSpecial(XlPasteType.xlPasteValues,
                    XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                copyWb.Save();
                copyWb.Close(defaultArg, defaultArg, defaultArg);

                Console.WriteLine(logPath + "_updated_salesreport_" + rep + ".txt");
                LogFile(logPath + "_updated_salesreport_" + rep + ".txt", "Succeefully updated Helias Sales Report.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(logPath + "_update_salesreport_" + rep + "_error.txt");
                LogFile(logPath + "_update_salesreport_" + rep + "_error.txt", ex.ToString());

                return;
            }
        }
        static void CopyToSalesRep(Application oXL, Workbook oWB, string template, string rep, string begCell, string endCell)
        {
            try
            {
                Object defaultArg = Type.Missing;
                Workbook wb = null;
                Worksheet ws = null;
                //string begCell = "E2";

                string date = now.ToString("MMMM dd, yyyy");

                string fileName = " - " + rep;

                wb = oXL.Workbooks.Open(path + "Helias Sales Report" + fileName + ".xlsx",
                            0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //ws = (Worksheet)wb.ActiveSheet;
                ws = (Worksheet)wb.Worksheets[template];

                ((Worksheet)oWB.Worksheets[template]).get_Range(begCell, endCell).Copy(Type.Missing);
                ws.get_Range(begCell, endCell).PasteSpecial(XlPasteType.xlPasteValues,
                    XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                wb.Save();
                wb.Close(defaultArg, defaultArg, defaultArg);

                Console.WriteLine(logPath + "_updated_salesreport_" + rep + ".txt");
                LogFile(logPath + "_updated_salesreport_" + rep + ".txt", "Succeefully updated Helias Sales Report.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(logPath + "_update_salesreport_" + rep + "_error.txt");
                LogFile(logPath + "_update_salesreport_" + rep + "_error.txt", ex.ToString());

                return;
            }
        }
        static decimal GetSalesByCustomer(string customer, IList<Sales> list)
        {
            decimal value = 0;
            var filteredList = list.Where(x => x.Customer.ToLower().Equals(customer.ToLower())).ToList();
            if (filteredList.Count() > 0)
                value = filteredList.Sum(s => s.Amount) ?? 0;

            return value;
        }

        static void LogFile(string file, string message)
        {
            if (!string.IsNullOrEmpty(file) || !String.IsNullOrEmpty(message))
            {
                // System.IO.File.WriteAllText(file, message);
                Console.WriteLine($"{message}");
            }
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
            List<Distributor> distributorList = new List<Distributor>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand com = new SqlCommand("SELECT Customer, WestRow, EastRow, CentralRow, Sheet FROM Distributor WHERE WestRow > 0 OR EastRow > 0 OR CentralRow > 0;", conn))
                {
                    SqlDataReader reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        distributorList.Add(new Distributor
                        {
                            Customer = reader["Customer"].ToString(),
                            WestRow = Convert.ToInt32(reader["WestRow"]),
                            EastRow = Convert.ToInt32(reader["EastRow"]),
                            CentralRow = Convert.ToInt32(reader["CentralRow"]),
                            Sheet = Convert.ToString(reader["Sheet"]),
                        });
                    }
                    reader.Close();
                }
                conn.Close();
            }
            return distributorList;
        }
    }
}
