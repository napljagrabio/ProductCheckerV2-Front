using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("product_checker_platforms")]
    public class Platform
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("name")]
        public string? Name { get; set; }

        [Column("domain")]
        public string? Domain { get; set; }

        [Column("availability")]
        public int Availability { get; set; }

        [Column("created_at")]
        public DateTime? CreatedAt { get; set; }

        [Column("updated_at")]
        public DateTime? UpdateAt { get; set; }

        public List<string>? Domains { get => Domain?.Split(',', StringSplitOptions.RemoveEmptyEntries)?.Select(s => s.Trim())?.ToList(); }

        public string? Image { get => $"{Name?.ToString().ToLower()}.png"; }
    }
}
