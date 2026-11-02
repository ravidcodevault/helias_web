using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OYA_Web_Application.Models.Services;
using OYA_Web_Application.Services.KeyVault_Service;
using OYA_Web_Application.Services.LDAP_Service;
using System.Reflection;
using System.Security.Claims;

namespace OYA_Web_Application.Controllers
{
	public class AuthenticationController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public const string SessionKeyName = "_Name";
        public AuthenticationController(IConfiguration configuration, ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
		{
			_logger = logger;
			_configuration = configuration;
            _webHostEnvironment = webHostEnvironment;
		}

		// GET: HomeController1
		public ActionResult Login()
		{
			return View();
		}


		[HttpPost]
		public async Task<ActionResult> Login(string username, string password)
		{
			if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
				return View();

            bool localDevelopmentLogin = _webHostEnvironment.IsDevelopment()
                && string.Equals(Request.Host.Host, "localhost", StringComparison.OrdinalIgnoreCase)
                && (Request.Host.Port == 5000 || Request.Host.Port == 5295)
                && HttpContext.Connection.RemoteIpAddress != null
                && System.Net.IPAddress.IsLoopback(HttpContext.Connection.RemoteIpAddress);
            string ldapServer = "";
            string supportUserName;
            string supportPassword;
            if (localDevelopmentLogin)
            {
                // Reuse the existing local test account without contacting Azure or LDAP.
                supportUserName = "imran@blueskippers.com";
                supportPassword = "P@s$W0rd#123";
            }
            else
            {
                AzureKeyVaultService service = new AzureKeyVaultService();
                ldapServer = service.RetrieveSecret("LDAP-Server");
                supportUserName = service.RetrieveSecret("SupportUsername");
                supportPassword = service.RetrieveSecret("SupportPassword");
            }
			

			LdapResponse ldapResponse = new LdapResponse();
			if (username == supportUserName && password == supportPassword)
			{
				ldapResponse = new LdapResponse
				{
					isAuthenticated = true,
					isAuthorizeduser = true,
					cn = "Admin",
					SAMAccountName = "Admin",
					UserPrincipalName = "Admin",
				};
			}
			else
			{
				if (localDevelopmentLogin)
                {
                    ViewBag.isAuthenticated = "Error";
                    return View();
                }
                LdapAuthenticationService ldapService = new LdapAuthenticationService();
				ldapResponse = ldapService.ValidateCredentials(ldapServer, username, password);
			}

			if (ldapResponse == null)
			{
				ViewBag.TehnicalError = true;
				return View();
			}

			if (ldapResponse != null && ldapResponse.isAuthenticated && ldapResponse.isAuthorizeduser)
			{
                HttpContext.Session.SetString(SessionKeyName, supportUserName);
                var claims = new List<Claim>
					{
						new Claim(ClaimTypes.Name, ldapResponse.cn),
						new Claim(ClaimTypes.NameIdentifier, ldapResponse.UserPrincipalName),
						new Claim("SAMAccountName", ldapResponse.SAMAccountName)
					};

				var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

				var authProperties = new AuthenticationProperties
				{
					IsPersistent = true,
					ExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(30)
				};

				await HttpContext.SignInAsync(
						CookieAuthenticationDefaults.AuthenticationScheme,
						new ClaimsPrincipal(claimsIdentity),
						authProperties);

				// Set a value in session
				HttpContext.Session.SetString("Username", ldapResponse.cn);

				ViewBag.isAuthenticated = "Success";
				return RedirectToAction("Dashboard", "Home");
			}
			else
			{
				ViewBag.isAuthenticated = ldapResponse != null && ldapResponse.isAuthenticated ? "" : "Error";
				ViewBag.isAuthorizeduser = ldapResponse != null && ldapResponse.isAuthorizeduser ? "" : "Error";
				return View();
			}
		}

		public async Task<IActionResult> Logout()
		{
			// Sign out the user
			await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
			HttpContext.Session.Clear();
			return RedirectToAction("Login", "Authentication");
		}
	}
}
