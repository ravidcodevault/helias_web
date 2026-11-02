using OYA_Web_Application.Models.Data.Dashboard;

namespace OYA_Web_Application.Models.ViewModel
{
	public class DashboardVM
	{
		public DashboardSales dashboardSales { get; set; } //= new DashboardSales();
		public List<CustomerSales> customerSales { get; set; } //= new CustomerSales();

		public List<MonthlySales> oyaMonthlySales { get; set; } // = new MonthlySales();
		public List<YearlySales> oyaYearlySales { get; set; } //= new MonthlySales();

		public List<MonthlySales> pinoMonthlySales { get; set; } //= new MonthlySales();
		public List<YearlySales> pinoYearlySales { get; set; } //= new MonthlySales();
	}
}
