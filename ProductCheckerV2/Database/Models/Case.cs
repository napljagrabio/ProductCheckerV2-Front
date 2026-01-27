using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("cases")]
    public class Case
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("case_number")]
        public string? CaseNumber { get; set; }

        [Column("deleted_at")]
        public DateTime? DeletedAt { get; set; }
    }
}

