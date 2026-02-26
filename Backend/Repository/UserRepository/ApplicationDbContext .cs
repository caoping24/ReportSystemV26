using CenterUser.Repository.Models;
using Microsoft.EntityFrameworkCore;

namespace CenterUser.Repository
{
    public class ApplicationDbContext : DbContext
    {
        public DbSet<User> User => Set<User>();

        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
               : base(options)
        {
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // todo：目前只用对单个表实体进行软删除过滤
            modelBuilder.Entity<User>().HasQueryFilter(p => !p.IsDelete);
        }
    }
}
