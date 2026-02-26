using CenterBackend.common;
using CenterBackend.Constant;
using CenterBackend.Dto;
using CenterBackend.Exceptions;
using CenterBackend.IUserServices;
using Masuit.Tools;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class UserController : ControllerBase
    {
        private readonly IUserService userService;

        public UserController(IUserService userService)
        {
            this.userService = userService;
        }

        /// <summary>
        /// 注册
        /// </summary>
        /// <param name="registerDto"></param>
        /// <returns></returns>
        /// <exception cref="BusinessException"></exception>
        [HttpPost("register")]
        public async Task<BaseResponse<long>> Register([FromBody] RegisterDto registerDto)
        {
            if (registerDto.UserAccount.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户账号不能为空");
            }
            if (registerDto.UserPassword.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户密码不能为空");
            }
            if (!registerDto.UserPassword.Equals(registerDto.CheckPassword))
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "两次密码不一致");
            }
            var result = await userService.UserRegister(registerDto);
            return ResultUtils<long>.Success(result);
        }

        /// <summary>
        /// 登陆
        /// </summary>
        /// <param name="loginDto"></param>
        /// <returns></returns>
        [HttpPost("login")]
        public async Task<BaseResponse<UserDto>> Login([FromBody] LoginDto loginDto)
        {
            if (loginDto.UserAccount.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户账号不能为空");
            }
            if (loginDto.UserPassword.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户密码不能为空");
            }
            var result = await userService.Login(loginDto);
            return ResultUtils<UserDto>.Success(result);
        }

        /// <summary>
        /// 注销
        /// </summary>
        /// <returns></returns>
        [HttpPost("logout")]
        public BaseResponse<int> UserLogout()
        {
            var result = userService.UserLogout();
            return ResultUtils<int>.Success(result);
        }

        /// <summary>
        /// 获取当前用户
        /// </summary>
        /// <returns></returns>
        [HttpGet("current")]
        public async Task<BaseResponse<UserDto>?> getCurrentUser()
        {
            var userObj = HttpContext.Session.GetString(UserConstant.USER_LOGIN_STATE);
            // todo：暂时返回空， 防止第一次加载前端那边报错，前端还未做全局异常处理
            if (userObj == null)
            {
                return null;
            }
            var user = JsonSerializer.Deserialize<UserDto>(userObj);
            if (user == null)
            {
                return null;
            }
            var result = await userService.GetSafetyUser(user.Id);
            return ResultUtils<UserDto>.Success(result);
        }

        /// <summary>
        /// 搜索用户
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        /// <exception cref="BusinessException"></exception>
        [HttpGet("search")]
        public async Task<BaseResponse<List<UserDto>>> SearchUsers(string userName = "")
        {
            // 仅管理员可以查询
            if (!isAdmin())
            {
                throw new BusinessException(ErrorCode.NO_AUTH, "缺少管理员权限");
            }
            var reuslt = await userService.SearchUsers(userName);
            return ResultUtils<List<UserDto>>.Success(reuslt);
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        /// <exception cref="BusinessException"></exception>
        [HttpPost("delete")]
        public async Task<BaseResponse<bool>> DeleteUser([FromBody] long id)
        {
            // 仅管理员可以删除
            if (!isAdmin())
            {
                throw new BusinessException(ErrorCode.NO_AUTH, "缺少管理员权限");
            }
            if (id <= 0)
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "id不能小于0");
            }
            var result = await userService.DeleteUser(id);
            return ResultUtils<bool>.Success(result);
        }

        /// <summary>
        /// 是否为管理员
        /// </summary>
        /// <param name="httpContent"></param>
        /// <returns></returns>
        private bool isAdmin()
        {
            // 仅管理员可查询
            var userObj = HttpContext.Session.GetString(UserConstant.USER_LOGIN_STATE);
            if (userObj == null)
            {
                return false;
            }
            var user = JsonSerializer.Deserialize<UserDto>(userObj);
            return user != null && user.Role == UserConstant.ADMIN_ROLE;
        }
    }
}
