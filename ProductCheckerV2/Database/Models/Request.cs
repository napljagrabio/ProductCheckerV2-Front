using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ProductCheckerV2.Database.Models
{
    [Table("requests")]
    public class Request
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("id")]
        public int Id { get; set; }

        [Column("request_info_id")]
        public int RequestInfoId { get; set; }

        [Column("status")]
        public RequestStatus Status { get; set; } = RequestStatus.PENDING;

        [Column("request_ended")]
        public DateTime? RequestEnded { get; set; }

        [Column("rescan_info_id")]
        public int RescanInfoId { get; set; }

        [Column("priority")]
        public int Priority { get; set; }

        [Column("created_at")]
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        [Column("updated_at")]
        public DateTime? UpdatedAt { get; set; }

        [Column("deleted_at")]
        public DateTime? DeletedAt { get; set; }

        [ForeignKey("RequestInfoId")]
        public virtual RequestInfo RequestInfo { get; set; }
    }

    public enum RequestStatus
    {
        PENDING,
        PROCESSING,
        SUCCESS,
        FAILED,
        COMPLETED_WITH_ISSUES
    }
}