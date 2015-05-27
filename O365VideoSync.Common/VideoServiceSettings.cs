using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365VideoSync.Common
{
    public class VideoServiceSettings
    {
        public string ChannelUrlTemplate { get; set; }
        public string IsVideoPortalEnabled { get; set; }
        public string PlayerUrlTemplate { get; set; }
        public string VideoPortalLayoutsUrl { get; set; }
        public string VideoPortalUrl { get; set; }
    }
}
