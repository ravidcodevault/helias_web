using System.Data.SqlClient;
using System.Data;
using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using GemBox.Spreadsheet.Drawing;
using Intuit.Ipp.Data;

namespace OYA_Web_Application.Services.ReportServices
{
    public class CommissionReport
    {
        private readonly IConfiguration _configuration;
        private readonly string path = "";
        private readonly string commissionTemplate = "";
        private readonly string connectionString = "";

        private static string CommissionCustomerNameCol = "A";
        private static string CommissionTypeCol = "B";
        private static string CommissionShipStateCol = "C";
        private static string CommissionDateCol = "D";
        private static string CommissionNumCol = "E";
        private static string CommissionCustomerCol = "F";
        private static string CommissionItemCol = "G";
        private static string CommissionItemDescriptionCol = "H";
        private static string CommissionQtyCol = "I";
        private static string CommissionSalesCol = "J";
        private static string CommissionAmountCol = "K";
        private static string CommissionFirstOrderCol = "L";
        private static string CommissionRateCol = "M";
        private static string CommissionCommissionCol = "N";
        private static string CommissionContractorCol = "O";
        private static string CommissionShipCityCol = "P";
        private static string CommissionTotalDenotion = "Q";
        private static int CommissionStartRow = 2;




        byte[] content = null;

        List<CommissionData> commissionDataList = new List<CommissionData>();
        List<ContractorCustomer> contractorCustomerList = new List<ContractorCustomer>();

        public CommissionReport(IConfiguration configuration)
        {
            _configuration = configuration;
            if (configuration != null)
            {
                path = _configuration["TemplatePath"];
                commissionTemplate = _configuration["FilePaths:CommissionReportTemplate"];
                connectionString = _configuration.GetConnectionString("DefaultConnection");
            }
        }
        private void FetchDetailsData(string contrator, string startDate, string endDate)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand("GetCommisionReport", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    connection.Open();

                    command.Parameters.Add("@contractor", SqlDbType.VarChar).Value = contrator;
                    command.Parameters.Add("@startDate", SqlDbType.VarChar).Value = startDate;
                    command.Parameters.Add("@endDate", SqlDbType.VarChar).Value = endDate;

                    SqlDataAdapter adapt = new SqlDataAdapter(command);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);

                    var commissionTable = ds.Tables[0];

                    commissionDataList.Clear();

                    foreach (DataRow _row in commissionTable.Rows)
                    {
                        var commissionData = new CommissionData();
                        commissionData.Type = Convert.ToString(_row["Type"]);
                        commissionData.ShipState = Convert.ToString(_row["ShipState"]);
                        commissionData.Date = Convert.ToDateTime(_row["Date"]);
                        commissionData.Num = Convert.ToString(_row["Num"]);
                        commissionData.Customer = Convert.ToString(_row["Customer"]);
                        commissionData.Item = Convert.ToString(_row["Item"]);
                        commissionData.ItemDescription = Convert.ToString(_row["ItemDescription"]);
                        commissionData.Qty = Convert.ToInt32(_row["Qty"]);
                        commissionData.SalesPrice = Convert.ToDouble(_row["SalesPrice"]);
                        commissionData.Amount = Convert.ToDouble(_row["Amount"]);
                        commissionData.ShipCity = Convert.ToString(_row["ShipCity"]);
                        commissionData.Rate = Convert.ToInt32(_row["Rate"]);
                        commissionData.Commission = Convert.ToDouble(_row["Commission"]);
                        commissionData.Contractor = Convert.ToString(_row["Contractor"]);
                        commissionData.ShipCity = Convert.ToString(_row["ShipCity"]);
                        //commissionData.FirstOrderDate = Convert.ToDateTime(_row["FirstOrderDate"]);

                        commissionDataList.Add(commissionData);
                    }

                    contractorCustomerList.Clear();

                    var CustomerTable = ds.Tables[1];
                    foreach (DataRow _row in CustomerTable.Rows)
                    {
                        contractorCustomerList.Add(new ContractorCustomer
                        {
                            Customer = Convert.ToString(_row["Customer"]),
                            FirstOrderDate = Convert.ToDateTime(_row["FirstOrderDate"])
                        });
                    }
                }
            }
        }
        private void GenerateDetailExcel()
        {

            var oWB = ExcelFile.Load(path + commissionTemplate);
            var oSheet = oWB.Worksheets[0];

            var index = CommissionStartRow;


            foreach (var customer in contractorCustomerList.Where(c => commissionDataList.Select(_c => _c.Customer).Contains(c.Customer)))
            {
                oSheet.Cells[CommissionCustomerNameCol + index].Value = customer.Customer;
                oSheet.Cells[CommissionFirstOrderCol + index].Value = customer.FirstOrderDate.ToString("yyyy-MM-dd");
                index++;

                var qtyTotal = 0;
                var amountTotal = 0.0;
                var commissionTotal = 0.0;
                foreach (var item in commissionDataList.Where(c => c.Customer == customer.Customer))
                {
                    oSheet.Cells[CommissionTypeCol + index].Value = item.Type;
                    oSheet.Cells[CommissionShipStateCol + index].Value = item.ShipState;
                    oSheet.Cells[CommissionDateCol + index].Value = item.Date.ToString("yyyy-MM-dd");
                    oSheet.Cells[CommissionNumCol + index].Value = item.Num;
                    oSheet.Cells[CommissionCustomerCol + index].Value = item.Customer;
                    oSheet.Cells[CommissionItemCol + index].Value = item.Item;
                    oSheet.Cells[CommissionItemDescriptionCol + index].Value = item.ItemDescription.Trim();
                    qtyTotal += item.Qty;
                    oSheet.Cells[CommissionQtyCol + index].Value = item.Qty;
                    oSheet.Cells[CommissionSalesCol + index].Value = Math.Round(item.SalesPrice, 2);
                    amountTotal += item.Amount;
                    oSheet.Cells[CommissionAmountCol + index].Value = item.Amount;
                    oSheet.Cells[CommissionRateCol + index].Value = item.Rate;
                    commissionTotal += item.Commission;
                    oSheet.Cells[CommissionCommissionCol + index].Value = item.Commission;
                    oSheet.Cells[CommissionContractorCol + index].Value = item.Contractor;
                    oSheet.Cells[CommissionShipCityCol + index].Value = item.ShipCity;
                    index++;
                }
                oSheet.Cells[CommissionQtyCol + index].Value = qtyTotal;
                oSheet.Cells[CommissionAmountCol + index].Value = amountTotal;
                oSheet.Cells[CommissionCommissionCol + index].Value = commissionTotal;
                oSheet.Cells[CommissionTotalDenotion + index].Value = "total";
                index++;
            }

            oSheet.Cells[CommissionQtyCol + index].Value = commissionDataList.Sum(i => i.Qty);
            oSheet.Cells[CommissionAmountCol + index].Value = commissionDataList.Sum(i => i.Amount);
            oSheet.Cells[CommissionCommissionCol + index].Value = commissionDataList.Sum(i => i.Commission);
            oSheet.Cells[CommissionTotalDenotion + index].Value = "total";

            using (var memoryStream = new MemoryStream())
            {
                oWB.Save(memoryStream, SaveOptions.XlsxDefault);

                content = memoryStream.ToArray();
            }

        }
        public byte[] DownloadExcel(string contrator = "02", DateTime? startDate = null, DateTime? endDate = null)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            FetchDetailsData(contrator, startDate.Value.ToString("MM/dd/yyyy"), endDate.Value.ToString("MM/dd/yyyy"));
            GenerateDetailExcel();

            return content;
        }
    }
    public class ContractorCustomer
    {
        public string Customer { get; set; }
        public DateTime FirstOrderDate { get; set; }
    }
    public class CommissionData
    {
        public string Type { get; set; }
        public string ShipState { get; set; }
        public DateTime Date { get; set; }
        public string Num { get; set; }
        public string Customer { get; set; }
        public string Item { get; set; }
        public string ItemDescription { get; set; }
        public int Qty { get; set; }
        public double SalesPrice { get; set; }
        public double Amount { get; set; }
        public DateTime FirstOrderDate { get; set; }
        public int Rate { get; set; }
        public double Commission { get; set; }
        public string Contractor { get; set; }
        public string ShipCity { get; set; }
    }
}
