using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365VideoSync.Common
{
    public class Video
    {
        public int OnPremItemId { get; set; }
        public string ChannelID { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Description { get; set; }
        public string DisplayFormUrl { get; set; }
        public string FileName { get; set; }
        public string OwnerName { get; set; }
        public string ServerRelativeUrl { get; set; }
        public string ThumbnailUrl { get; set; }
        public string Title { get; set; }
        public string ID { get; set; }// This is the ID of the video in O365. Cant rename , otherwise newtonsof toObject will break
        public string Url { get; set; }
        public double VideoDurationInSeconds { get; set; }
        public double VideoProcessingStatus { get; set; }
        public double ViewCount { get; set; }
        public string YammerObjectUrl { get; set; }
    }
}
