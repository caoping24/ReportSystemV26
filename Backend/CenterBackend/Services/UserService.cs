using CenterBackend.common;
using CenterBackend.Constant;
using CenterBackend.Dto;
using CenterBackend.Exceptions;
using CenterBackend.IUserServices;
using CenterUser.Repository;
using CenterUser.Repository.Models;
using Mapster;
using Masuit.Tools;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;

namespace CenterBackend.Services
{
    public class UserService : IUserService
    {
        private readonly IRepository<User> userRepository;
        private readonly IUnitOfWork unitOfWork;
        private readonly IHttpContextAccessor httpContextAccessor;

        public UserService(IRepository<User> userRepository, IUnitOfWork unitOfWork, IHttpContextAccessor httpContextAccessor)
        {
            this.userRepository = userRepository;
            this.unitOfWork = unitOfWork;
            this.httpContextAccessor = httpContextAccessor;
        }

        public async Task<bool> DeleteUser(long id)
        {
            await userRepository.DeleteByIdAsync(id);
            await unitOfWork.SaveChangesAsync();
            return true;
        }

        public async Task<long> UserRegister(RegisterDto registerDto)
        {
            // 1.校验
            if (registerDto.UserAccount.IsNullOrEmpty() || registerDto.UserPassword.IsNullOrEmpty() || registerDto.CheckPassword.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "参数为空");
            }
            if (await userRepository.db.SingleOrDefaultAsync(s => s.UserAccount == registerDto.UserAccount) != null)
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户名已经存在");
            }
            if (!registerDto.UserPassword.Equals(registerDto.CheckPassword))
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "两次输入的密码不一致");
            }

            // 2.添加用户
            var user = new User();
            user.UserAccount = registerDto.UserAccount;
            // todo: 密码暂时未加密
            user.UserPassword = registerDto.UserPassword;
            user.UserStatus = 0;
            //默认注册的是普通用户
            user.Role = 0;
            await userRepository.AddAsync(user);
            await unitOfWork.SaveChangesAsync();
            return user.Id;
        }

        public async Task<UserDto> Login(LoginDto loginDto)
        {
            // 1.校验
            if (loginDto.UserAccount.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户账号不能为空");
            }
            if (loginDto.UserPassword.IsNullOrEmpty())
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "用户密码不能为空");
            }

            // 查询用户是否存在
            var user = await userRepository.db.SingleOrDefaultAsync(s => s.UserAccount == loginDto.UserAccount && s.UserPassword == loginDto.UserPassword);
            // 用户不存在
            if (user == null)
            {
                throw new BusinessException(ErrorCode.NULL_ERROR, "用户不存在");
            }
            // 2.用户信息脱敏
            var safetyUser = user.Adapt<UserDto>();
            // 3.记录用户登录态
            if (httpContextAccessor.HttpContext != null)
            {
                var session = httpContextAccessor.HttpContext.Session;
                session.SetString(UserConstant.USER_LOGIN_STATE, JsonSerializer.Serialize(safetyUser));
            }
            return safetyUser;
        }

        public async Task<UserDto> GetSafetyUser(long id)
        {
            var user = await userRepository.GetByIdAsync(id);
            return user.Adapt<UserDto>();
        }

        public async Task<List<UserDto>> SearchUsers(string userName)
        {
            var result = await userRepository.db.WhereIf(!userName.IsNullOrEmpty(), s => s.UserName.Contains(userName)).ToListAsync();
            return result.Adapt<List<UserDto>>();
        }

        public int UserLogout()
        {
            if (httpContextAccessor.HttpContext != null)
            {
                var session = httpContextAccessor.HttpContext.Session;
                session.Remove(UserConstant.USER_LOGIN_STATE);
            }
            return 1;
        }
    }
}
