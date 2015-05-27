using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using O365VideoSync.Common;
using Newtonsoft.Json.Linq;
using SPO = Microsoft.SharePoint.Client;
namespace O365VideoSync.O365Utilities
{
    public class O365Utilities
    {
        public string siteUrl { get; set; }
        public string userID { get; set; }
        public string password { get; set; }
        private SPO.SharePointOnlineCredentials _spoCredentials;
        private SPO.SharePointOnlineCredentials SpoCredentials
        {
            get
            {
                if (_spoCredentials != null)
                {
                    return _spoCredentials;
                }
                var securePassword = new System.Security.SecureString();
                foreach (var c in password)
                {
                    securePassword.AppendChar(c);
                }
                _spoCredentials = new SPO.SharePointOnlineCredentials(userID, securePassword);
                return _spoCredentials;
            }
        }
        public VideoServiceSettings  VideoServiceSettings{
        get {
            var endpointUri = siteUrl + "/_api/VideoService.discover";
            var t = MakeRestCall(endpointUri, SpoCredentials);
            var settings = t["d"];
            VideoServiceSettings vs = settings.ToObject<VideoServiceSettings>();
            return vs;

    }}
        public List<VideoChannel> Channels
        {
            get
            {
                var endpointUri = VideoServiceSettings.VideoPortalUrl + "/_api/VideoService/Channels";
                var t = MakeRestCall(endpointUri, SpoCredentials);

                var d = t["d"];
                var results = d["results"];
                var returnVal = results.ToObject<List<VideoChannel>>();
                return returnVal;

            }
        }
        public  List<Video> GetVideos(string ChannelId)
        {
            var endpointUri = string.Format("{0}/_api/VideoService/Channels('{1}')/Videos", VideoServiceSettings.VideoPortalUrl, ChannelId);
            var t = MakeRestCall(endpointUri, SpoCredentials);
            var d = t["d"];
            var resulkts = d["results"];
            var returnVal = resulkts.ToObject<List<Video>>();
            return returnVal;
        }

       

        public  VideoChannel GetChannelByName(string ChannelTitle)
        {
            var endpointUri = VideoServiceSettings.VideoPortalUrl + "/_api/VideoService/Channels";
            var t = MakeRestCall(endpointUri, SpoCredentials);

            var d = t["d"];
            var results = d["results"];
            var channels = results.ToObject<List<VideoChannel>>();
            var selectedChannel = channels.Find(channel => channel.Title == ChannelTitle);

            return selectedChannel;

        }
     
        private static JToken MakeRestCall(string endpointUri, ICredentials credentials)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var result = client.DownloadString(endpointUri);
                var t = Newtonsoft.Json.Linq.JToken.Parse(result);
                return t;

            }
        }
    }
}
