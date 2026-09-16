using Microsoft.EntityFrameworkCore;
using OutlookAddinProject.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddHttpClient();
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

// ==== CONFIGURE THIS: point to your real database ====
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// CORS is kept here only in case something OTHER than this project's own task pane
// needs to call the API (e.g. a separate admin portal on another domain).
// The add-in itself doesn't need this, since taskpane.html and the API are
// now served from the exact same origin.
builder.Services.AddCors(options =>
{
    options.AddPolicy("DefaultPolicy", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

// Serves everything in wwwroot/ (manifest.xml, taskpane.html, commands.html,
// commands.js, assets/*.png) as static files, at the site's root.
app.UseDefaultFiles();
app.UseStaticFiles();

app.UseCors("DefaultPolicy");
app.UseSession();

app.UseAuthorization();

app.MapControllers();
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Login}/{action=Index}/{id?}");

app.Run();
