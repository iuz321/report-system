using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// ? 註冊 GoogleSheetService（這行最重要）
builder.Services.AddSingleton<GoogleSheetService>();

// 加入 MVC
builder.Services.AddControllersWithViews();

var app = builder.Build();

// 錯誤處理
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

// 路由（預設直接進 Report）
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Report}/{action=Index}/{id?}");

app.Run();