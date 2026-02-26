namespace CenterBackend.common
{
    public class ResultUtils<T>
    {
        // 成功响应的静态方法，放在ResultUtils类中
        public static BaseResponse<T> Success(T data)
        {
            return new BaseResponse<T>(0, data, "ok");
        }

        public static BaseResponse<T> error()
        {
            return new BaseResponse<T>(500, default(T), "查询失败");
        }
    }
}
