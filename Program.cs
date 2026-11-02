using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Office.Interop.Excel;
using OYA_Web_Application.Hubs;

var builder = WebApplication.CreateBuilder(args);

// Persist cookie keys outside versioned deployment folders. DPAPI binds the keys
// to this Windows account; app identity remains stable across folder changes.
var protection = builder.Services.AddDataProtection()
    .SetApplicationName("OYA.Web.Application");
// IIS uses the framework default key storage for its worker identity.
// Interactive Windows launches keep the existing per-user protected key store.
var hostedInIis = string.Equals(System.Diagnostics.Process.GetCurrentProcess().ProcessName,
    "w3wp", StringComparison.OrdinalIgnoreCase)
    || !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ASPNETCORE_IIS_PHYSICAL_PATH"))
    || !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ASPNETCORE_TOKEN"));
if (OperatingSystem.IsWindows() && !hostedInIis)
{
    var stateDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OyaWebApplication", "DataProtection-Keys");
    Directory.CreateDirectory(stateDirectory);
    protection.PersistKeysToFileSystem(new DirectoryInfo(stateDirectory))
        .ProtectKeysWithDpapi();
}

// Add services to the container.
builder.Services.AddSignalR();
builder.Services.AddCors(options => {
    options.AddPolicy("CORSPolicy", builder => builder.AllowAnyMethod().AllowAnyHeader().AllowCredentials().SetIsOriginAllowed((hosts) => true));
});
builder.Services.AddControllersWithViews();

// Add session services
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
	options.IdleTimeout = TimeSpan.FromMinutes(20);
	options.Cookie.Name = "OYA.Web.Session";
	options.Cookie.HttpOnly = true;
	options.Cookie.IsEssential = true;
});
// Add authentication services
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
		.AddCookie(options =>
		{
			options.Cookie.Name = "OYA.Web.Authentication";
			options.LoginPath = "/Authentication/Login"; // Redirect to login page
			options.ExpireTimeSpan = TimeSpan.FromMinutes(20);
        });
builder.Services.TryAddSingleton<IHttpContextAccessor, HttpContextAccessor>();
//builder.Services.AddControllersWithViews().AddRazorRuntimeCompilation();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
	app.UseExceptionHandler("/Home/Error");
}
app.UseStaticFiles();
app.UseCors("CORSPolicy");
app.UseRouting();

app.UseSession();
app.UseAuthentication();
app.UseAuthorization();

app.MapHub<MessageHub>("/messagehub");

app.MapControllerRoute(
	name: "default",
	pattern: "{controller=Home}/{action=Dashboard}/{id?}");

app.Run();
