using System.Data.SqlClient;
using System.Data;
using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using GemBox.Spreadsheet.Drawing;
using Intuit.Ipp.Data;

namespace OYA_Web_Application.Services.ReportServices
{
	public class AgingReport
	{
		private readonly IConfiguration _configuration;
		private readonly string path = "";
		private readonly string summaryTemplate = "";
		private readonly string detailTemplate = "";
		private readonly string connectionString = "";

		static string SummaryCutomerCell = "A";
		static string SummaryCurrentCell = "C";
		static string Summary1_30Cell = "E";
		static string Summary31_60Cell = "G";
		static string Summary61_90Cell = "I";
		static string Summary90PlusCell = "K";
		static string SummaryTotalCell = "M";
		static string SummaryTotalRowMarking = "N";
		static int SummaryStartRow = 2;

		static string DetailCategoryCell = "A";
		static string DetailTypeCell = "C";
		static string DetailDateCell = "E";
		static string DetailRefNumberCell = "G";
		static string DetailCustomerCell = "K";
		static string DetailTermsCell = "M";
		static string DetailDueDateCell = "O";
		static string DetailAgingCell = "R";
		static string DetailOpenBalanceCell = "T";
		static int DetailtartRow = 2;



		byte[] content = null;

		List<AgingSummary> agingSummaryList = new List<AgingSummary>();

		public AgingReport(IConfiguration configuration)
		{
			_configuration = configuration;
			if (configuration != null)
			{
				path = _configuration["TemplatePath"];
				summaryTemplate = _configuration["FilePaths:AgingSummaryReportTemplate"];
				detailTemplate = _configuration["FilePaths:AgingDetailsReportTemplate"];
				connectionString = _configuration.GetConnectionString("DefaultConnection");
			}
		}
		private void FetchSummaryData(string type)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				using (SqlCommand command = new SqlCommand("GetAgingReportSummary", connection))
				{
					command.CommandType = CommandType.StoredProcedure;
					connection.Open();

					command.Parameters.Add("@reportType", SqlDbType.VarChar).Value = "Summary";
                    command.Parameters.Add("@reportFor", SqlDbType.VarChar).Value = type;

                    SqlDataAdapter adapt = new SqlDataAdapter(command);
					DataSet ds = new DataSet();
					adapt.Fill(ds);

					var thisMonthDataOya = ds.Tables[0];

					agingSummaryList.Clear();

					foreach (DataRow _row in thisMonthDataOya.Rows)
					{
						var agingSummary = new AgingSummary();
						//agingSummary.TxnType = Convert.ToString(_row["TxnType"]);
						agingSummary.CustomerName = Convert.ToString(_row["Customer"]);
						agingSummary.Amount = Convert.ToDouble(_row["Amount"]);
						agingSummary.Category = Convert.ToString(_row["Category"]);
						agingSummary.AccountRefFullName = Convert.ToString(_row["AccountRefFullName"]);

						agingSummaryList.Add(agingSummary);
					}
				}
			}
		}
		private void GenerateSummaryExcel(string type = "OYA")
		{
			var oWB = ExcelFile.Load(path + summaryTemplate);
			var oSheet = oWB.Worksheets[0];

			var index = SummaryStartRow;

			var customerLists = agingSummaryList.Where(r => r.AccountRefFullName == type).DistinctBy(x => x.CustomerName).Select(c => c.CustomerName).ToList();

			var totalCurrent = 0.0;
			var total1_30 = 0.0;
			var total31_60 = 0.0;
			var total61_90 = 0.0;
			var total90Plus = 0.0;

			foreach (var customer in customerLists)
			{
				oSheet.Cells[SummaryCutomerCell + index].Value = customer;

				var current = agingSummaryList.Where(l => l.CustomerName == customer && l.Category == "0").Sum(a => a.Amount);
				var current1_30 = agingSummaryList.Where(l => l.CustomerName == customer && l.Category == "1-30").Sum(a => a.Amount);
				var current1_60 = agingSummaryList.Where(l => l.CustomerName == customer && l.Category == "31-60").Sum(a => a.Amount);
				var current1_90 = agingSummaryList.Where(l => l.CustomerName == customer && l.Category == "61-90").Sum(a => a.Amount);
				var current90Plus = agingSummaryList.Where(l => l.CustomerName == customer && l.Category == "90+").Sum(a => a.Amount);

				oSheet.Cells[SummaryCurrentCell + index].Value = current;
				oSheet.Cells[Summary1_30Cell + index].Value = current1_30;
				oSheet.Cells[Summary31_60Cell + index].Value = current1_60;
				oSheet.Cells[Summary61_90Cell + index].Value = current1_90;
				oSheet.Cells[Summary90PlusCell + index].Value = current90Plus;

				totalCurrent += current;
				total1_30 += current1_30;
				total31_60 += current1_60;
				total61_90 += current1_90;
				total90Plus += current90Plus;

				index++;
			}

			oSheet.Cells[SummaryCurrentCell + index].Value = totalCurrent;
			oSheet.Cells[Summary1_30Cell + index].Value = total1_30;
			oSheet.Cells[Summary31_60Cell + index].Value = total31_60;
			oSheet.Cells[Summary61_90Cell + index].Value = total61_90;
			oSheet.Cells[Summary90PlusCell + index].Value = total90Plus;

			oSheet.Cells[SummaryTotalRowMarking + index].Value = "total";

			using (var memoryStream = new MemoryStream())
			{
				oWB.Save(memoryStream, SaveOptions.XlsxDefault);

				content = memoryStream.ToArray();
			}
		}

		private void FetchDetailsData(string type)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				using (SqlCommand command = new SqlCommand("GetAgingReportSummary", connection))
				{
					command.CommandType = CommandType.StoredProcedure;
					connection.Open();

					command.Parameters.Add("@reportType", SqlDbType.VarChar).Value = "Detail";
					command.Parameters.Add("@reportFor", SqlDbType.VarChar).Value = type;

					SqlDataAdapter adapt = new SqlDataAdapter(command);
					DataSet ds = new DataSet();
					adapt.Fill(ds);

					var thisMonthDataOya = ds.Tables[0];

					agingSummaryList.Clear();

					foreach (DataRow _row in thisMonthDataOya.Rows)
					{
						var agingSummary = new AgingSummary();
						agingSummary.TxnType = Convert.ToString(_row["TxnType"]);
						agingSummary.CustomerName = Convert.ToString(_row["Customer"]);
						agingSummary.Amount = Convert.ToDouble(_row["Amount"]);
						agingSummary.Category = Convert.ToString(_row["Category"]);
						agingSummary.RefNumber = Convert.ToString(_row["RefNumber"]);
						agingSummary.ClassRefFullName = Convert.ToString(_row["ClassRefFullName"]);
						agingSummary.SalesRepRefFullName = Convert.ToString(_row["SalesRepRefFullName"]);
						agingSummary.CustomerTypeRefFullName = Convert.ToString(_row["CustomerTypeRefFullName"]);
						agingSummary.TxnDate = Convert.ToDateTime(_row["TxnDate"]);
						agingSummary.DueDate = Convert.ToDateTime(_row["DueDate"]);
						agingSummary.Year = Convert.ToInt32(_row["Year"]);
						agingSummary.DiffInDays = Convert.ToInt32(_row["DiffInDays"]);
						agingSummary.TermsRefFullName = Convert.ToString(_row["TermsRefFullName"]);
						agingSummary.AccountRefFullName = Convert.ToString(_row["AccountRefFullName"]);

						agingSummaryList.Add(agingSummary);
					}
				}
			}
		}
		private void GenerateDetailExcel(string type = "OYA")
		{

			var oWB = ExcelFile.Load(path + detailTemplate);
			var oSheet = oWB.Worksheets[0];

			var index = DetailtartRow;
			oSheet.Cells[DetailCategoryCell + index].Value = "Current";
			index++;
			foreach (var item in agingSummaryList.Where(r => r.Category == "0" && r.AccountRefFullName == type))
			{
				oSheet.Cells[DetailTypeCell + index].Value = item.TxnType;
				oSheet.Cells[DetailDateCell + index].Value = item.TxnDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailRefNumberCell + index].Value = item.RefNumber;
				oSheet.Cells[DetailCustomerCell + index].Value = item.CustomerName;
				oSheet.Cells[DetailTermsCell + index].Value = item.TermsRefFullName;
				oSheet.Cells[DetailDueDateCell + index].Value = item.DueDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailAgingCell + index].Value = "0";
				oSheet.Cells[DetailOpenBalanceCell + index].Value = item.Amount;
				index++;
			}

			oSheet.Cells[DetailCategoryCell + index].Value = "Total Current";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Where(r => r.Category == "0").Sum(o => o.Amount);
			index++;

			oSheet.Cells[DetailCategoryCell + index].Value = "1 - 30";
			index++;

			foreach (var item in agingSummaryList.Where(r => r.Category == "1-30"))
			{

				oSheet.Cells[DetailTypeCell + index].Value = item.TxnType;
				oSheet.Cells[DetailDateCell + index].Value = item.TxnDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailRefNumberCell + index].Value = item.RefNumber;
				oSheet.Cells[DetailCustomerCell + index].Value = item.CustomerName;
				oSheet.Cells[DetailTermsCell + index].Value = item.TermsRefFullName;
				oSheet.Cells[DetailDueDateCell + index].Value = item.DueDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailAgingCell + index].Value = "1";
				oSheet.Cells[DetailOpenBalanceCell + index].Value = item.Amount;
				index++;
			}
			oSheet.Cells[DetailCategoryCell + index].Value = "Total 1 - 30";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Where(r => r.Category == "1-30").Sum(o => o.Amount);
			index++;

			oSheet.Cells[DetailCategoryCell + index].Value = "31 - 60";
			index++;

			foreach (var item in agingSummaryList.Where(r => r.Category == "31-60"))
			{

				oSheet.Cells[DetailTypeCell + index].Value = item.TxnType;
				oSheet.Cells[DetailDateCell + index].Value = item.TxnDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailRefNumberCell + index].Value = item.RefNumber;
				oSheet.Cells[DetailCustomerCell + index].Value = item.CustomerName;
				oSheet.Cells[DetailTermsCell + index].Value = item.TermsRefFullName;
				oSheet.Cells[DetailDueDateCell + index].Value = item.DueDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailAgingCell + index].Value = "2";
				oSheet.Cells[DetailOpenBalanceCell + index].Value = item.Amount;
				index++;
			}
			oSheet.Cells[DetailCategoryCell + index].Value = "Total 31-60";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Where(r => r.Category == "31-60").Sum(o => o.Amount);
			index++;

			oSheet.Cells[DetailCategoryCell + index].Value = "61 - 90";
			index++;

			foreach (var item in agingSummaryList.Where(r => r.Category == "61-90"))
			{

				oSheet.Cells[DetailTypeCell + index].Value = item.TxnType;
				oSheet.Cells[DetailDateCell + index].Value = item.TxnDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailRefNumberCell + index].Value = item.RefNumber;
				oSheet.Cells[DetailCustomerCell + index].Value = item.CustomerName;
				oSheet.Cells[DetailTermsCell + index].Value = item.TermsRefFullName;
				oSheet.Cells[DetailDueDateCell + index].Value = item.DueDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailAgingCell + index].Value = "3";
				oSheet.Cells[DetailOpenBalanceCell + index].Value = item.Amount;
				index++;
			}
			oSheet.Cells[DetailCategoryCell + index].Value = "Total 61 - 90";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Where(r => r.Category == "61-90").Sum(o => o.Amount);
			index++;
			oSheet.Cells[DetailCategoryCell + index].Value = "> 90";
			index++;

			foreach (var item in agingSummaryList.Where(r => r.Category == "90+"))
			{
				oSheet.Cells[DetailTypeCell + index].Value = item.TxnType;
				oSheet.Cells[DetailDateCell + index].Value = item.TxnDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailRefNumberCell + index].Value = item.RefNumber;
				oSheet.Cells[DetailCustomerCell + index].Value = item.CustomerName;
				oSheet.Cells[DetailTermsCell + index].Value = item.TermsRefFullName;
				oSheet.Cells[DetailDueDateCell + index].Value = item.DueDate.ToString("MM/dd/yyyy");
				oSheet.Cells[DetailAgingCell + index].Value = "4";
				oSheet.Cells[DetailOpenBalanceCell + index].Value = item.Amount;
				index++;
			}
			oSheet.Cells[DetailCategoryCell + index].Value = "Total > 90";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Where(r => r.Category == "90+").Sum(o => o.Amount);
			index++;
			oSheet.Cells[DetailCategoryCell + index].Value = "TOTAL";
			oSheet.Cells[DetailOpenBalanceCell + index].Value = agingSummaryList.Sum(o => o.Amount);

			using (var memoryStream = new MemoryStream())
			{
				oWB.Save(memoryStream, SaveOptions.XlsxDefault);

				content = memoryStream.ToArray();
			}

		}
		public byte[] DownloadExcel(string type = "OYA", string report = "summary")
		{
			SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
			switch (report.ToLower())
			{
				case "details":
					FetchDetailsData(type);
					GenerateDetailExcel(type);
					break;
				case "summary":
				default:
					FetchSummaryData(type);
					GenerateSummaryExcel(type);
					break;
			}

			return content;
		}
	}
	public class AgingSummary
	{
		public string TxnType { get; set; }
		public string CustomerName { get; set; }
		public double Amount { get; set; }
		public string Category { get; set; }
		public string RefNumber { get; set; }
		public string ClassRefFullName { get; set; }
		public string AccountRefFullName { get; set; }
		public string SalesRepRefFullName { get; set; }
		public string CustomerTypeRefFullName { get; set; }
		public DateTime TxnDate { get; set; }
		public DateTime DueDate { get; set; }
		public int Year { get; set; }
		public int DiffInDays { get; set; }
		public string TermsRefFullName { get; set; }
	}
}
