using Microsoft.EntityFrameworkCore;
using ProductCheckerV2.Database.Models;

namespace ProductCheckerV2.Database
{
    public class ProductCheckerDbContext : DbContext
    {
        public DbSet<Requests> Requests { get; set; }
        public DbSet<RequestInfos> RequestInfos { get; set; }
        public DbSet<ProductListings> ProductListings { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var connectionString = Common.ConfigurationManager.GetConnectionString();

                optionsBuilder.UseMySql(connectionString,
                    ServerVersion.AutoDetect(connectionString));
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Requests>()
                .Property(r => r.Status)
                .HasConversion<string>();

            modelBuilder.Entity<Requests>()
                .HasOne(r => r.RequestInfo)
                .WithMany(ri => ri.Requests)
                .HasForeignKey(r => r.RequestInfoId)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<ProductListings>()
                .HasOne(p => p.RequestInfo)
                .WithMany()
                .HasForeignKey(p => p.RequestInfoId)
                .OnDelete(DeleteBehavior.Cascade);
        }
    }
}