using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("campaigns")]
    public class Campaign
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("name")]
        public string? Name { get; set; }

        [Column("status")]
        public int Status { get; set; }

        [Column("deleted_at")]
        public DateTime? DeletedAt { get; set; }
    }
}

