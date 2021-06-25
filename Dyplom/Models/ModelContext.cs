using Dyplom.Models;
using System.Data.Entity;

namespace Dyplom.Models
{
    class ModelContext : DbContext
    {
        public ModelContext() : base("DefaultConnection")
        {
        }

        public DbSet<User> Users { get; set; }
        public DbSet<Students> Students { get; set; }
        public DbSet<LeadTeachers> LeadTeachers { get; set; }
        public DbSet<Management> Management { get; set; }

        public DbSet<Classes> Classes { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Management>().ToTable("Management");
        }
    }
}
