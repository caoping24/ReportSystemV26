namespace CenterBackend.common
{
    /// <summary>
    /// 通用返回对象
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class BaseResponse<T>
    {
        public int code { get; set; }

        public T data { get; set; }

        public string message { get; set; }

        public string? description { get; set; }

        public BaseResponse(int code, T data, string message = "")
        {
            this.code = code;
            this.data = data;
            this.message = message;
        }
    }
}
