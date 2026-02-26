using CenterBackend.IServices;
using CenterBackend.IUserServices;
using CenterBackend.Logging;
using CenterBackend.Middlewares;
using CenterBackend.Services;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Services;
using CenterUser.Repository;
using Microsoft.AspNetCore.Session;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.OpenApi.Models;
namespace CenterBackend
{
    public class Program
    {
        // 保持原有 Main 行为，单独运行时不变
        public static void Main(string[] args)
        {
            var app = BuildWebApplication(args);
            app.Run();
        }

        // 对外提供的工厂方法：构建 WebApplication（但不 Run）
        // contentRootPath 可用于在外部（如 WPF）指定静态文件所在的目录
        public static WebApplication BuildWebApplication(string[]? args = null, string? contentRootPath = null, int port = 5260)
        {
            var builder = WebApplication.CreateBuilder(args ?? Array.Empty<string>());

            if (!string.IsNullOrEmpty(contentRootPath))
            {
                // 指定 ContentRoot（确保 wwwroot 可被找到）
                builder.Environment.ContentRootPath = contentRootPath;
            }

            var configuration = builder.Configuration;

            builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));
            builder.Services.AddScoped(typeof(IReportRepository<>), typeof(ReportRepository<>));
            builder.Services.AddScoped(typeof(IReportRecordRepository<>), typeof(ReportRecordRepository<>));

            string defaultConnection = configuration.GetConnectionString("DefaultConnection") ?? string.Empty;
            builder.Services.AddDbContext<ApplicationDbContext>(options => options.UseSqlServer(defaultConnection));
            builder.Services.AddDbContext<CenterReportDbContext>(options => options.UseSqlServer(defaultConnection));

            builder.Services.AddScoped<IUnitOfWork, UnitOfWork>();
            builder.Services.AddScoped<IReportUnitOfWork, ReportUnitOfWork>();

            builder.Services.AddScoped<IUserService, UserService>();
            builder.Services.AddScoped<IReportService, ReportService>();
            builder.Services.AddScoped<IFileServices, FileService>();
            builder.Services.AddScoped<IReportRecordService, ReportRecordService>();
            builder.Services.AddScoped<IDashboardService, DashboardService>();
            // 注册日志服务（单例），FileLogger 会使用 IWebHostEnvironment.ContentRootPath 定位到 wwwroot/log
            builder.Services.AddSingleton<IAppLogger, FileLogger>();

            // 显式注册控制器所在的程序集，确保在 ReportServer 进程内也能发现控制器
            builder.Services.AddControllers()
                .AddApplicationPart(typeof(Program).Assembly) // 确保包含 CenterBackend 的控制器
                .AddControllersAsServices();

            builder.Services.AddSpaStaticFiles(spaConfig =>
            {
                spaConfig.RootPath = "wwwroot/dist";
            });
            builder.Services.RemoveAll<ISessionStore>();
            builder.Services.RemoveAll<IDistributedCache>();
            builder.Services.AddDistributedMemoryCache();
            builder.Services.AddSession(options =>
            {
                options.IdleTimeout = TimeSpan.FromMinutes(30);
                options.Cookie.HttpOnly = true;
                options.Cookie.IsEssential = true;
                options.Cookie.SameSite = SameSiteMode.None;
                options.Cookie.SecurePolicy = CookieSecurePolicy.None;
                options.Cookie.Name = "ReportSystem_SessionId";
            });

            var allowedOrigins = configuration["CorsPolicy:AllowedOrigins"]?.Split(',') ?? Array.Empty<string>();
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("Policy", policy =>
                {
                    policy.WithOrigins(allowedOrigins)
                          .AllowAnyHeader()
                          .AllowAnyMethod()
                          .AllowCredentials()
                          .WithExposedHeaders("Content-Disposition");// 标准：暴露非简单响应头，前端才能读取 Content-Disposition
                });
            });
            builder.Services.AddAuthentication("CookieAuth")
                .AddCookie("CookieAuth", options =>
                {
                    options.Cookie.Name = "ReportSystem_SessionId";
                    options.Events.OnRedirectToLogin = context =>
                    {
                        context.Response.StatusCode = StatusCodes.Status401Unauthorized;
                        return Task.CompletedTask;
                    };
                    options.Events.OnRedirectToAccessDenied = context =>
                    {
                        context.Response.StatusCode = StatusCodes.Status403Forbidden;
                        return Task.CompletedTask;
                    };
                    options.ExpireTimeSpan = TimeSpan.FromMinutes(20);
                    options.SlidingExpiration = true;

                    options.Cookie.SameSite = SameSiteMode.None;
                    options.Cookie.SecurePolicy = CookieSecurePolicy.None;
                });
            builder.Services.AddHttpContextAccessor();
            builder.Services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "报表系统API", Version = "v1" });
            });

            // Kestrel 绑定到 loopback（本机），避免 ListenAnyIP 导致防火墙弹窗
            builder.WebHost.UseKestrel(options =>
            {
                // 【核心修正】从 ListenLocalhost 改为 ListenAnyIP，允许局域网访问
                options.ListenAnyIP(port); // 替换原 options.ListenLocalhost(port);
                options.Limits.MaxConcurrentConnections = 1000;
                options.AllowSynchronousIO = true;
                options.Limits.MaxConcurrentUpgradedConnections = 1000;
            });

            var app = builder.Build();

            // 临时请求日志（便于确认请求是否到达服务器；可上线前移除）
            app.Use(async (ctx, next) =>
            {
                Console.WriteLine($"[REQ] {ctx.Request.Method} {ctx.Request.Path}");
                await next();
                Console.WriteLine($"[RES] {ctx.Request.Method} {ctx.Request.Path} -> {ctx.Response.StatusCode}");
            });

            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "报表系统API v1");
                    c.RoutePrefix = "swagger";
                });
            }

            app.UseSpaStaticFiles();
            app.UseMiddleware<GlobalExceptionMiddleware>();
            app.UseCors("Policy");
            app.UseSession();
            app.UseRouting();
            app.UseAuthentication();
            app.UseAuthorization();
            app.MapControllers();
            app.MapFallbackToFile("dist/index.html");

            // 返回构建好的 app，调用方负责 StartAsync / StopAsync / Run
            return app;
        }
    }
}