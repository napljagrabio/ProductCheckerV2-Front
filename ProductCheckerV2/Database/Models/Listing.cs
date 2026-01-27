using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("listings")]
    public class Listing
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        [Column("campaign_id")]
        public int CampaignId { get; set; }

        [Column("case_id")]
        public int CaseId { get; set; }

        [Column("platform_id")]
        public int PlatformId { get; set; }

        [Column("qflag_id")]
        public int QfalgId { get; set; }

        [Column("url")]
        public string? Url { get; set; }

        [Column("deleted_at")]
        public DateTime? DeletedAt { get; set; }
    }
}

