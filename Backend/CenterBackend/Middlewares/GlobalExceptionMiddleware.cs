using CenterBackend.common;
using CenterBackend.Exceptions;
using Masuit.Tools.Systems;
using System.Net;
using System.Text.Json;

namespace CenterBackend.Middlewares
{
    /// <summary>
    /// 全局异常处理中间件
    /// </summary>
    public class GlobalExceptionMiddleware
    {
        // 这是一个委托，指向下一个中间件
        private readonly RequestDelegate _next;

        public GlobalExceptionMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        // 这是中间件的主要执行方法，接收一个 HttpContext 参数。
        public async Task InvokeAsync(HttpContext context)
        {
            try
            {
                // 该方法尝试执行下一个中间件(_next(context))，并在发生异常时捕获这些异常。
                await _next(context);
            }
            catch (BusinessException ex)
            {
                await HandleBusinessExceptionAsync(context, ex);
            }
            catch (Exception)
            {
                // 对于非业务异常，我们只返回通用的错误消息和500状态码，不暴露详细信息
                await HandleOtherExceptionAsync(context);
            }
        }

        private Task HandleBusinessExceptionAsync(HttpContext context, BusinessException exception)
        {
            context.Response.ContentType = "application/json";
            context.Response.StatusCode = exception.Code;

            return context.Response.WriteAsync(JsonSerializer.Serialize(new
            {
                exception.Code,
                exception.Message,
                exception.Description
            }));
        }

        private Task HandleOtherExceptionAsync(HttpContext context)
        {
            context.Response.ContentType = "application/json";
            context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;

            return context.Response.WriteAsync(JsonSerializer.Serialize(new
            {
                Code = (int)HttpStatusCode.InternalServerError,
                Message = ErrorCode.SYSTEM_ERROR.GetDescription(),
            }));
        }
    }
}
