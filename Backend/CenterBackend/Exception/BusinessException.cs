using CenterBackend.common;
using Masuit.Tools.Systems;

namespace CenterBackend.Exceptions
{
    /// <summary>
    /// 自定义异常
    /// </summary>
    public class BusinessException : Exception
    {
        /**
         * 异常码
         */
        public int Code { get; }

        /**
         * 描述
         */
        public string Description { get; }

        public BusinessException(ErrorCode errorCode)
            : base(errorCode.GetDescription())
        {
            Code = (int)errorCode;
        }

        public BusinessException(ErrorCode errorCode, string description)
            : base(errorCode.GetDescription())
        {
            Code = (int)errorCode;
            Description = description;
        }
    }

}
