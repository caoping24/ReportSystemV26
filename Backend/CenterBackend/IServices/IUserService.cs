using CenterBackend.Dto;

namespace CenterBackend.IUserServices
{
    public interface IUserService
    {
        Task<long> UserRegister(RegisterDto registerDto);

        Task<UserDto> Login(LoginDto loginDto);

        Task<List<UserDto>> SearchUsers(string userName);

        Task<bool> DeleteUser(long id);

        Task<UserDto> GetSafetyUser(long id);

        int UserLogout();
    }
}
