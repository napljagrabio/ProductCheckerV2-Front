using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("api_endpoints")]
    public class ApiEndpoint
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("key")]
        public string? Key { get; set; }

        [Column("value")]
        public string? Value { get; set; }

        [Column("created_at")]
        public DateTime? CreatedAt { get; set; }
    }
}