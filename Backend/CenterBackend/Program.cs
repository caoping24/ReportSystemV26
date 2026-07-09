using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Middlewares;
using CenterBackend.Services;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Services;
using CenterUser.Repository;
using Hangfire;
using Hangfire.Common;
using Hangfire.SqlServer;
using Microsoft.AspNetCore.Session;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.OpenApi.Models;

namespace CenterBackend
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var app = await BuildWebApplicationAsync(args);
            app.Run();
        }

        // 对外提供的工厂方法：构建 WebApplication(但不 Run)
        // contentRootPath 可用于在外部(如 WPF)指定静态文件所在的目录
        public static async Task<WebApplication> BuildWebApplicationAsync(string[]? args = null, string? contentRootPath = null, int port = 5260)
        {
            var builder = WebApplication.CreateBuilder(args ?? Array.Empty<string>());

            if (!string.IsNullOrEmpty(contentRootPath))
            {
                // 指定 ContentRoot(确保 wwwroot 可被找到)
                builder.Environment.ContentRootPath = contentRootPath;
            }

            var configuration = builder.Configuration;

            builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));
            builder.Services.AddScoped(typeof(IReportRepository<>), typeof(ReportRepository<>));
            builder.Services.AddScoped(typeof(IReportRecordRepository<>), typeof(ReportRecordRepository<>));
            builder.Services.AddScoped(typeof(IOperatorInputDataRepository<>), typeof(OperatorInputDataRepository<>));
            string defaultConnection = configuration.GetConnectionString("DefaultConnection") ?? string.Empty;
            builder.Services.AddDbContext<ApplicationDbContext>(options => options.UseSqlServer(defaultConnection));
            builder.Services.AddDbContext<CenterReportDbContext>(options => options.UseSqlServer(defaultConnection));

            // 配置 Hangfire，使用 SQL Server 存储任务数据
            builder.Services.AddHangfire(cfg =>
            {
                cfg.UseSimpleAssemblyNameTypeSerializer()
                   .UseRecommendedSerializerSettings()
                   .UseSqlServerStorage(defaultConnection, new SqlServerStorageOptions
                   {
                       PrepareSchemaIfNecessary = true,
                       SchemaName = "hangfire"
                   });
            });

            builder.Services.AddHangfireServer();

            // 你的任务服务(确保已注册)
            builder.Services.AddScoped<IBackGroundServices, BackGroundServices>();

            builder.Services.AddScoped<IUnitOfWork, UnitOfWork>();
            builder.Services.AddScoped<IReportUnitOfWork, ReportUnitOfWork>();

            builder.Services.AddScoped<IDashboardService, DashboardService>();
            builder.Services.AddScoped<IDataToViewService, DataToViewService>();
            builder.Services.AddScoped<IDataViewToExcel, DataViewToExcel>();
            builder.Services.AddScoped<IFileServices, FileService>();
            builder.Services.AddScoped<IReportService, ReportService>();
            builder.Services.AddScoped<IReportRecordService, ReportRecordService>();
            builder.Services.AddScoped<IUserService, UserService>();

            // 注册日志服务(单例)，FileLogger 会使用 IWebHostEnvironment.ContentRootPath 定位到 wwwroot/log
            builder.Services.AddSingleton<IAppLogger, FileLogger>();

            // 2026年7月8日新增 —— 筛选配置服务
            builder.Services.AddSingleton<IFilterConfigService, FilterConfigService>();
            builder.Services.AddScoped<DataFilterService>();

            // 显式注册控制器所在的程序集，确保在 ReportServer 进程内也能发现控制器
            builder.Services.AddControllers()
                .AddApplicationPart(typeof(Program).Assembly)
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
                options.Cookie.SameSite = SameSiteMode.Lax;
            });

            // CORS
            var corsPolicy = configuration.GetSection("CorsPolicy");
            var allowedOrigins = corsPolicy.GetValue<string>("AllowedOrigins") ?? "";
            var origins = allowedOrigins
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct()
                .ToArray();
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("Policy", policy =>
                {
                    if (origins.Length > 0)
                        policy.WithOrigins(origins)
                              .AllowAnyHeader()
                              .AllowAnyMethod()
                              .AllowCredentials();
                    else
                        policy.AllowAnyOrigin()
                              .AllowAnyHeader()
                              .AllowAnyMethod();
                });
            });

            // 添加响应压缩
            builder.Services.AddResponseCompression(options =>
            {
                options.EnableForHttps = true;
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

            // Kestrel 绑定到 loopback(本机)，避免 ListenAnyIP 导致防火墙弹窗
            builder.WebHost.UseKestrel(options =>
            {
                options.ListenAnyIP(port);
                options.Limits.MaxConcurrentConnections = 1000;
                options.AllowSynchronousIO = true;
                options.Limits.MaxConcurrentUpgradedConnections = 1000;
            });

            var app = builder.Build();

            // 加载筛选配置
            using (var initScope = app.Services.CreateScope())
            {
                var configService = initScope.ServiceProvider.GetRequiredService<IFilterConfigService>();
                await configService.ReloadAsync();
            }

            // 启动定时任务
            var tz = TimeZoneInfo.FindSystemTimeZoneById("China Standard Time");
            var opt = new RecurringJobOptions { TimeZone = tz };

            using (var scope = app.Services.CreateScope())
            {
                var mgr = scope.ServiceProvider.GetRequiredService<IRecurringJobManager>();

                mgr.AddOrUpdate(
                    "daily-0810",
                    Job.FromExpression<IBackGroundServices>(x => x.Daily0810()),
                    "10 8 * * *",
                    opt);

                mgr.AddOrUpdate(
                    "weekly-mon-0820",
                    Job.FromExpression<IBackGroundServices>(x => x.WeeklyMon0820()),
                    "20 8 * * 1",
                    opt);

                mgr.AddOrUpdate(
                    "monthly-1st-0830",
                    Job.FromExpression<IBackGroundServices>(x => x.MonthlyDay1_0830()),
                    "30 8 1 * *",
                    opt);
            }

            // 临时请求日志
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
                    c.HeadContent = "<a href='/' style='position:absolute;top:14px;left:16px;color:white;'>返回首页</a>";
                });
            }
            string useHangfireDashboard = configuration.GetValue<string>("UseHangfireDashboard:ON") ?? string.Empty;
            if (useHangfireDashboard == "true")
            {
                app.UseHangfireDashboard("/hangfire/main");
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

            return app;
        }
    }
}
