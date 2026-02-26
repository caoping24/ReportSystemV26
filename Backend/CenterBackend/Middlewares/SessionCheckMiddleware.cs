using CenterBackend.common;
using CenterBackend.Constant;
using Masuit.Tools.Systems;
using System.Net;
using System.Text.Json;

namespace CenterBackend.Middlewares
{
    /// <summary>
    /// 全局Session过期检查中间件
    /// </summary>
    public class SessionCheckMiddleware
    {
        private readonly RequestDelegate _next;
        // 无需检查Session的接口路径（根据你的业务调整，排除公开接口）
        private readonly List<string> _excludePaths = new()
        {
            "/api/user/login",    // 登录接口（替换为你实际的登录接口路径）
            "/api/user/register", // 注册接口（如有）
            "/swagger",           // Swagger文档（开发环境保留）
            "/health",            // 健康检查接口（如有）
            "/api/file/upload",// 示例：无需鉴权的文件上传接口（根据实际调整）
            "/api/user/current"
        };

        public SessionCheckMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            // 1. 获取请求路径并统一转小写，避免大小写问题
            var requestPath = context.Request.Path.ToString().ToLower();

            // 2. 判断是否为无需检查的接口
            bool isExcluded = _excludePaths.Any(path => requestPath.StartsWith(path.ToLower()));

            if (!isExcluded)
            {
                // 3. 检查Session中是否存在有效用户标识（核心：UserId为Session存储的用户唯一标识）

                var userObj = context.Session.GetString(UserConstant.USER_LOGIN_STATE);
                if (userObj == null)
                {
                    // 4. Session过期/未登录：返回401响应（与你的异常返回格式保持一致）
                    context.Response.ContentType = "application/json";
                    context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;

                    // 适配你现有异常返回格式（Code+Message）
                    var errorResponse = new
                    {
                        Code = (int)HttpStatusCode.Unauthorized,
                        Message = ErrorCode.SESSION_EXPIRED.GetDescription(), // 建议在ErrorCode中新增SESSION_EXPIRED枚举
                        Description = "会话已过期，请重新登录"
                    };

                    var json = JsonSerializer.Serialize(errorResponse);
                    await context.Response.WriteAsync(json);
                    return; // 终止请求管道，不执行后续中间件  
                }

            }

            // 5. Session有效/无需检查：继续执行下一个中间件
            await _next(context);
        }
    }

    /// <summary>
    /// 中间件扩展方法（方便注册）
    /// </summary>
    public static class SessionCheckMiddlewareExtensions
    {
        public static IApplicationBuilder UseSessionCheck(this IApplicationBuilder app)
        {
            return app.UseMiddleware<SessionCheckMiddleware>();
        }
    }
}