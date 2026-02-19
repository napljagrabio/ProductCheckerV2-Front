using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("product_checker_listings")]
    public class ProductListing
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("id")]
        public long Id { get; set; }

        [Column("request_info_id")]
        public long RequestInfoId { get; set; }

        [Column("listing_id")]
        public long ListingId { get; set; }

        [Column("case_number")]
        [MaxLength(255)]
        public string? CaseNumber { get; set; }

        [Column("platform")]
        [MaxLength(255)]
        public string? Platform { get; set; }

        [Column("url")]
        [MaxLength(255)]
        public string? Url { get; set; }

        [Column("status")]
        [MaxLength(50)]
        public string? UrlStatus { get; set; }

        [Column("checked_date")]
        [MaxLength(50)]
        public string? CheckedDate { get; set; }

        [Column("error_detail")]
        public string? ErrorDetail { get; set; }

        [Column("note")]
        [MaxLength(255)]
        public string? Note { get; set; }

        [Column("created_at")]
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        [ForeignKey("RequestInfoId")]
        public virtual RequestInfo RequestInfo { get; set; }
    }
}
