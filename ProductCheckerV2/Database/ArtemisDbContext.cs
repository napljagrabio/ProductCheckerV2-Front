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
    internal class ArtemisDbContext : DbContext
    {
        public DbSet<Campaign> Campaigns { get; set; }
        public DbSet<Case> Cases { get; set; }
        public DbSet<Listing> Listings { get; set; }
        public DbSet<Platform> Platforms { get; set; }
        public DbSet<Qflag> Qflag { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var connectionString = Common.ConfigurationManager.GetConnectionString("ArtemisDbContext");

                optionsBuilder.UseMySql(connectionString,
                    ServerVersion.AutoDetect(connectionString));
            }
        }
    }
}