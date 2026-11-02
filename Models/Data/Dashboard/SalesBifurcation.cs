using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using OYA_Web_Application.Models.ViewModel;

namespace OYA_Web_Application.Models.Data.Dashboard
{
	public class DashboardSales
	{
		public decimal TotalAmountPublicDirect { get; set; }
		public decimal TotalAmountForeign { get; set; }
		public decimal TotalAmountOtherThisMonth { get; set; }
		public decimal TotalAmountOverallOya { get; set; }
		public decimal OYA_This_Month_Sales { get; set; }
		public decimal Pino_This_Month_Sales { get; set; }
		public decimal SalonSales { get; set; }
	}
	public class SalesBifurcation
	{
		private readonly IConfiguration _configuration;
		string connectionString = "";
		public SalesBifurcation(IConfiguration configuration)
		{
			_configuration = configuration;
			connectionString = _configuration.GetConnectionString("DefaultConnection");
		}

		public DashboardSales GetOverallDashboardSales()
		{
			DashboardSales sales = new DashboardSales();

			try
			{
				using (SqlConnection connection = new SqlConnection(connectionString))
				{
					using (SqlCommand command = new SqlCommand("dbo.USP_GetOverallDashboardSales_V2", connection))
					{
						command.CommandType = CommandType.StoredProcedure;
						connection.Open();

						// Add parameters
						command.Parameters.Add("@thisYear", SqlDbType.Int).Value = DateTime.Now.Year;
						command.Parameters.Add("@thisMonth", SqlDbType.Int).Value = DateTime.Now.Month;

						using (SqlDataReader reader = command.ExecuteReader())
						{
							while (reader.Read())
							{
								//DashboardSales sales = new DashboardSales();
								sales.TotalAmountPublicDirect = reader["PublicSales"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["PublicSales"]);
								sales.TotalAmountForeign = reader["InternetOrder"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["InternetOrder"]);
								sales.TotalAmountOtherThisMonth = reader["OtherThisMonth"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["OtherThisMonth"]);
								sales.TotalAmountOverallOya = reader["OverallSales"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["OverallSales"]);
								sales.OYA_This_Month_Sales = reader["OYASales"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["OYASales"]);
								sales.Pino_This_Month_Sales = reader["PinoSales"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["PinoSales"]);
								sales.SalonSales = reader["SalonSales"] == DBNull.Value ? 0m : Convert.ToDecimal(reader["SalonSales"]);
								//salesList.Add(sales);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: " + ex.Message);
			}

			return sales;
		}

		public decimal GetPublicDirectSales()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
            SELECT
                SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                Customer ON Invoice.CustomerRefListID = Customer.ListID
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Invoice.TxnDate BETWEEN '2024-05-01' AND '2025-05-07'
                AND Item.FullName LIKE 'OY%'
                AND Invoice.SalesRepRefFullName IS NOT NULL
                AND Invoice.SalesRepRefFullName IN (SELECT Contractor FROM Contractor)";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public decimal GetForeignBusiness()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
            SELECT
                SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                Customer ON Invoice.CustomerRefListID = Customer.ListID
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Invoice.TxnDate BETWEEN '2024-05-01' AND '2025-05-07'
                AND Item.FullName LIKE 'OY%'
                AND Invoice.ClassRefFullName = 'Internet Order'";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public decimal GetOtherThisMonthSales()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
            SELECT
               SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                Customer ON Invoice.CustomerRefListID = Customer.ListID
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Invoice.TxnDate BETWEEN '2024-05-01' AND '2025-05-07'
                AND Customer.FullName NOT IN (
                    'OYA ISCS:ISCS Derek Moser', 
                    'OYA ISCS:ISCS Michelle Fanning',
                    'OYA M Beauty:Andrea''s', 
                    'OYA M Beauty:Ariel Derrick (Residence)', 
                    'OYA M Beauty:Breanna Taylor (Residence)',
                    'OYA M Beauty:Elite Salon', 
                    'OYA M Beauty:Flaunt Salon', 
                    'OYA M Beauty:Fran Medina (Residence)', 
                    'OYA M Beauty:GAL Studios', 
                    'OYA M Beauty:Irene Fernandez', 
                    'OYA M Beauty:Jagged Edge', 
                    'OYA M Beauty:Karen Poulsen', 
                    'OYA M Beauty:Koolfx Salon at Mesa Royale',              
                    'OYA M Beauty:Linda Reese (Residence)', 
                    'OYA M Beauty:Marc Reid Salon @ Steven Paul Salon',
                    'OYA M Beauty:Mira Bella Hair Salon',
                    'OYA M Beauty:R Salon Day Spa - Donna Raffanti',
                    'OYA M Beauty:Salon Boutique - Arrowhead', 
                    'OYA M Beauty:Schwahn Salon',
                    'OYA M Beauty:Shaggy''s', 
                    'OYA M Beauty:Shear Creations',
                    'OYA M Beauty:Stasyss Salon',
                    'OYA M Beauty:Sublime Salon',
                    'OYA Unique Salon Concepts:Lisa Gagne'
                )
                AND Customer.FullName NOT IN (
                    'OYA Uberbeauty808',
                    'OYA C.K. ColorHouse',
                    'OYA Store Professional',
                    'OYA Store Consumer',
                    'OYA M Beauty',
                    'OYA Josephs College',
                    'OYA Blaze Salon',
                    'OYA ProSalon Distribution Group',
                    'OYA Central States Supply',
                    'OYA Elite Salon Distributors USA OHIO',
                    'OYA Discover Salon Services',
                    'OYA Mid America Beauty Supply',
                    'OYA Economy',
                    'OYA Elite Salon Services',
                    'OYA Valley Beauty Supply',
                    'OYA Unique Salon Concepts',
                    'OYA Style Care',
                    'OYA ISCS',
                    'OYA Prime Salon Services',
                    'OYA Aidan Laura Aine Ltd - ALA',
                    'OYA Desbo Distribution Inc.',
                    'OYA SOID Distributors, P.R.',
                    'OYA De Bleu Beauty Supply',
                    'OYA Binani Desbo Group',
                    'OYA Braliza LLC',
                    'OYA NCN Beauty LLC',
                    'OYA NCN Beauty LLC Venezuela',
                    'OYA NCN Beauty LLC Ecuador',
                    'OYA Brothers Beauty Products'
                )
                 AND Item.FullName LIKE 'OY%'";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public decimal GetOverallOyaSales()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
            SELECT
                SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                Customer ON Invoice.CustomerRefListID = Customer.ListID
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Invoice.TxnDate BETWEEN '2024-05-01' AND '2024-05-07'
                AND Item.FullName LIKE 'OY%'";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public decimal GetOyaThisMonth()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
            SELECT
                SUM(InvoiceLine.Amount) AS OYA_This_Month_Sales
            FROM
                Invoice
            JOIN
                Customer ON Invoice.CustomerRefListID = Customer.ListID
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Invoice.TxnDate BETWEEN '2024-05-01' AND '2025-05-07'
                AND Customer.FullName IN (
                    'OYA Uberbeauty808',
                    'OYA C.K. ColorHouse',
                    'OYA Store Professional',
                    'OYA Store Consumer',
                    'OYA M Beauty',
                    'OYA Josephs College',
                    'OYA Blaze Salon',
                    'OYA ProSalon Distribution Group',
                    'OYA Central States Supply',
                    'OYA Elite Salon Distributors USA OHIO',
                    'OYA Discover Salon Services',
                    'OYA Mid America Beauty Supply',
                    'OYA Economy',
                    'OYA Elite Salon Services',
                    'OYA Valley Beauty Supply',
                    'OYA Unique Salon Concepts',
                    'OYA Style Care',
                    'OYA ISCS',
                    'OYA Prime Salon Services',
                    'OYA Aidan Laura Aine Ltd - ALA',
                    'OYA Desbo Distribution Inc.',
                    'OYA SOID Distributors, P.R.',
                    'OYA De Bleu Beauty Supply',
                    'OYA Binani Desbo Group',
                    'OYA Braliza LLC',
                    'OYA NCN Beauty LLC',
                    'OYA NCN Beauty LLC Venezuela',
                    'OYA NCN Beauty LLC Ecuador',
                    'OYA Brothers Beauty Products'
                )
                 AND Item.FullName LIKE 'OY%'";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public decimal GetPinoThisMonth()
		{
			decimal totalAmount = 0;

			// SQL query to execute
			string query = @"
				SELECT
				SUM(InvoiceLine.Amount) AS TotalAmount
				FROM
				Invoice
				JOIN
				Customer ON Invoice.CustomerRefListID = Customer.ListID
				JOIN
				InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
				JOIN
					Item ON InvoiceLine.ItemRefListID = Item.ListID
				WHERE
					Invoice.TxnDate BETWEEN '2024-05-01' AND '2024-05-07'
				AND Item.FullName LIKE 'PN%'";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				if (reader.Read())
				{
					// Checking if the value is not DBNull and then converting it to decimal
					if (!reader.IsDBNull(0))
					{
						totalAmount = reader.GetDecimal(0);
					}
				}

				reader.Close();
			}

			return totalAmount;
		}

		public List<CustomerSales> GetTopCustomerSales()
		{
			List<CustomerSales> topCustomerSales = new List<CustomerSales>();

			// SQL query to execute
			string query = @"
				SELECT TOP 8
					Customer.FullName AS CustomerName,
					SUM(InvoiceLine.Amount) AS TotalAmount
				FROM
					Invoice
				JOIN
					Customer ON Invoice.CustomerRefListID = Customer.ListID
				JOIN
					InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
				JOIN
					Item ON InvoiceLine.ItemRefListID = Item.ListID
				WHERE
					YEAR(Invoice.TxnDate) = YEAR(GETDATE())
					--AND Item.FullName LIKE 'PN%'
				GROUP BY Customer.FullName
				ORDER BY TotalAmount DESC";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				while (reader.Read())
				{
					CustomerSales customerSales = new CustomerSales();
					customerSales.CustomerName = reader.GetString(0);
					customerSales.TotalAmount = reader.GetDecimal(1);
					topCustomerSales.Add(customerSales);
				}

				reader.Close();
			}

			return topCustomerSales;
		}

		public List<MonthlySales> GetOYAMonthlySales()
		{
			List<MonthlySales> monthlySales = new List<MonthlySales>();

			// SQL query to execute
			string query = @"
                SELECT
                    MONTH(Invoice.TxnDate) AS MonthNumber,
                    DATENAME(MONTH, Invoice.TxnDate) AS Month,
                    SUM(InvoiceLine.Amount) AS TotalAmount
                FROM
                    Invoice
                JOIN
                    InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
                JOIN
                    Item ON InvoiceLine.ItemRefListID = Item.ListID
                WHERE
                    YEAR(Invoice.TxnDate) = YEAR(GETDATE())
                    AND Item.FullName LIKE 'OY%'
                GROUP BY 
                    MONTH(Invoice.TxnDate),
                    DATENAME(MONTH, Invoice.TxnDate)
            ORDER BY 
                MonthNumber";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				while (reader.Read())
				{
					MonthlySales monthlySale = new MonthlySales();
					monthlySale.MonthNumber = reader.GetInt32(0);
					monthlySale.Month = reader.GetString(1);
					monthlySale.TotalAmount = reader.GetDecimal(2);
					monthlySales.Add(monthlySale);
				}

				reader.Close();
			}

			return monthlySales;
		}

		public List<MonthlySales> GetPinoMonthlySales()
		{
			List<MonthlySales> monthlySales = new List<MonthlySales>();

			// SQL query to execute
			string query = @"
                SELECT
                    MONTH(Invoice.TxnDate) AS MonthNumber,
                    DATENAME(MONTH, Invoice.TxnDate) AS Month,
                    SUM(InvoiceLine.Amount) AS TotalAmount
                FROM
                    Invoice
                JOIN
                    InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
                JOIN
                    Item ON InvoiceLine.ItemRefListID = Item.ListID
                WHERE
                    YEAR(Invoice.TxnDate) = YEAR(GETDATE())
                    AND Item.FullName LIKE 'PN%'
                GROUP BY 
                    MONTH(Invoice.TxnDate),
                    DATENAME(MONTH, Invoice.TxnDate)
				ORDER BY MonthNumber
            ";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				while (reader.Read())
				{
					MonthlySales monthlySale = new MonthlySales();
					monthlySale.MonthNumber = reader.GetInt32(0);
					monthlySale.Month = reader.GetString(1);
					monthlySale.TotalAmount = reader.GetDecimal(2);
					monthlySales.Add(monthlySale);
				}

				reader.Close();
			}

			return monthlySales;
		}

		public List<YearlySales> GetOYAYearlySales()
		{
			List<YearlySales> yearlySales = new List<YearlySales>();

			// SQL query to execute
			string query = @"
            SELECT
                YEAR(Invoice.TxnDate) AS Year,
                SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Item.FullName LIKE 'OY%' AND YEAR(Invoice.TxnDate) > 2015
            GROUP BY 
                YEAR(Invoice.TxnDate)
            ORDER BY 
                Year";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				while (reader.Read())
				{
					YearlySales yearlySale = new YearlySales();
					yearlySale.Year = reader.GetInt32(0);
					yearlySale.TotalAmount = reader.GetDecimal(1);
					yearlySales.Add(yearlySale);
				}

				reader.Close();
			}

			return yearlySales;
		}

		public List<YearlySales> GetPinoYearlySales()
		{
			List<YearlySales> yearlySales = new List<YearlySales>();

			// SQL query to execute
			string query = @"
            SELECT
                YEAR(Invoice.TxnDate) AS Year,
                SUM(InvoiceLine.Amount) AS TotalAmount
            FROM
                Invoice
            JOIN
                InvoiceLine ON Invoice.TxnID = InvoiceLine.TxnID
            JOIN
                Item ON InvoiceLine.ItemRefListID = Item.ListID
            WHERE
                Item.FullName LIKE 'PN%' AND YEAR(Invoice.TxnDate) > 2015
            GROUP BY 
                YEAR(Invoice.TxnDate)
            ORDER BY 
                Year";

			// Establishing connection and executing the query
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(query, connection);
				connection.Open();
				SqlDataReader reader = command.ExecuteReader();

				// Reading the result
				while (reader.Read())
				{
					YearlySales yearlySale = new YearlySales();
					yearlySale.Year = reader.GetInt32(0);
					yearlySale.TotalAmount = reader.GetDecimal(1);
					yearlySales.Add(yearlySale);
				}

				reader.Close();
			}

			return yearlySales;
		}

	}

	public class YearlySales
	{
		public int Year { get; set; }
		public decimal TotalAmount { get; set; }
	}

	public class CustomerSales
	{
		public string CustomerName { get; set; }
		public decimal TotalAmount { get; set; }
	}

	public class MonthlySales
	{
		public int MonthNumber { get; set; }
		public string Month { get; set; }
		public decimal TotalAmount { get; set; }
	}
}
