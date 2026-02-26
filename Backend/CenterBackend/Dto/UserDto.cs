namespace CenterBackend.Dto
{
    public class UserDto
    {

        public long Id { get; set; }

        public string? UserName { get; set; }

        public string? UserAccount { get; set; }

        public string? AvatarUrl { get; set; }

        public int? Gender { get; set; }

        public string? Phone { get; set; }

        public string? Email { get; set; }

        /// <summary>
        /// 用户角色， 0= 普通， 1=管理员
        /// </summary>
        public int Role { get; set; }

    }
}
