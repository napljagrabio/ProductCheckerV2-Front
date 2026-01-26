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
        public DbSet<Platform> Platforms { get; set; }
        public DbSet<Port> Ports { get; set; }
        public DbSet<ApiEndpoint> ApiEndpoints { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var connectionString = Common.ConfigurationManager.GetConnectionString("ProductCheckerDbContext");

                optionsBuilder.UseMySql(connectionString,
                    ServerVersion.AutoDetect(connectionString));
            }
        }
    }
}