using System.ComponentModel;

namespace CenterBackend.common
{
    [Flags]
    public enum ErrorCode
    {
        [Description("请求参数错误")]
        PARAMS_ERROR = 40000,

        [Description("请求数据为空")]
        NULL_ERROR = 40001,

        [Description("未登录")]
        NOT_LOGIN = 40100,

        [Description("无权限")]
        NO_AUTH = 40101,

        [Description("系统内部异常")]
        SYSTEM_ERROR = 50000,

        [Description("session过期")]
        SESSION_EXPIRED = 401,
    }

}
