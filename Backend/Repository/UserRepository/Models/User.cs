using System.ComponentModel.DataAnnotations.Schema;

namespace CenterUser.Repository.Models
{
    [Table("User")] //指定表名称
    public class User : ISoftDelete
    {
        public long Id { get; set; }

        public string? UserName { get; set; }

        public string? UserAccount { get; set; }

        public string? AvatarUrl { get; set; }

        public int Gender { get; set; }

        public string? UserPassword { get; set; }

        public string? Phone { get; set; }

        public string? Email { get; set; }

        public int UserStatus { get; set; }

        public int Role { get; set; }

        public bool IsDelete { get; set; } = false;

    }
}
