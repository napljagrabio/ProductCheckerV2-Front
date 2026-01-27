using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("qflag")]
    public class Qflag
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("label")]
        public string? Label { get; set; }

        [Column("status")]
        public int Status { get; set; }

        [Column("deleted_at")]
        public DateTime? DeletedAt { get; set; }
    }
}

