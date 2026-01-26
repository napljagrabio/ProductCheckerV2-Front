using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("product_checker_ports")]
    public class Port
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("api")]
        public string? Api { get; set; }

        [Column("username")]
        public string? Username { get; set; }

        [Column("status")]
        public int Status { get; set; }

        [Column("created_at")]
        public DateTime? CreatedAt { get; set; }

        [Column("updated_at")]
        public DateTime? UpdatedAt { get; set; }
    }
}