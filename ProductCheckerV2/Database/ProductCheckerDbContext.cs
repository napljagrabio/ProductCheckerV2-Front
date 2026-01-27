using ProductCheckerV2.Database.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ProductCheckerV2.Database
{
    internal class ProductCheckerDbContext : DbContext
    {
        public DbSet<CrawlerPlatform> CrawlerPlatforms { get; set; }
        public DbSet<Port> Ports { get; set; }
        public DbSet<ApiEndpoint> ApiEndpoints { get; set; }
        public DbSet<Request> Requests { get; set; }
        public DbSet<RequestInfo> RequestInfos { get; set; }
        public DbSet<ProductListing> ProductListings { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var connectionString = Common.ConfigurationManager.GetConnectionString("ProductCheckerDbContext");

                optionsBuilder.UseMySql(connectionString,
                    ServerVersion.AutoDetect(connectionString));
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Request>()
                .Property(r => r.Status)
                .HasConversion<string>();

            modelBuilder.Entity<Request>()
                .HasOne(r => r.RequestInfo)
                .WithMany(ri => ri.Requests)
                .HasForeignKey(r => r.RequestInfoId)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<ProductListing>()
                .HasOne(p => p.RequestInfo)
                .WithMany()
                .HasForeignKey(p => p.RequestInfoId)
                .OnDelete(DeleteBehavior.Cascade);
        }
    }
}