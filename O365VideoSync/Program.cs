using System;
using System.Collections.Generic;
using System.Linq;

using System.Text;
using System.Threading.Tasks;
using SPO = Microsoft.SharePoint.Client;
using O365VideoSync.Common;
using O365VideoSync.O365Utilities;
using O365VideoSync.OnPremUtilities;
namespace O365VideoSync
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("****O365VideoSync started at " + DateTime.Now + " *****");
            var settings = Properties.Settings.Default;
     
            O365Utilities.O365Utilities o365Utils = new O365Utilities.O365Utilities()
            {
                siteUrl = settings.O365SharepointUrl,
                userID = settings.O365UserName,
                password = settings.O365Password
            };
            var videoServiceSettings = o365Utils.VideoServiceSettings;
            VideoChannel channel = o365Utils.GetChannelByName(settings.O365VideoChanleName);
            List<Video> o365Videos = o365Utils.GetVideos(channel.Id);

            OnPremUtilities.OnPremUtilities onPremUtils = new OnPremUtilities.OnPremUtilities()
            {
                siteUrl = settings.OnPremSharepointUrl,
                listTitle = settings.OnPremListName,
                userID = settings.OnPremUserName,
                password = settings.OnPremPassword
            };
            if (!onPremUtils.ListExists(settings.OnPremListName))
            {
                onPremUtils.CreateList(settings.OnPremListName, settings.ListDescription, settings.OnQuickLaunch);
            }
            List<Video> onPremVideos = onPremUtils.GetVideos();

            foreach (Video onPremVideo in onPremVideos)
            {
                var o365Version = o365Videos.Find(o3v => o3v.ID == onPremVideo.ID);// find an entry oin the o365 list with the same videoid
                if (o365Version == null)
                {
                    onPremUtils.RemoveVideo(onPremVideo);
                }
                else
                {
                    if (
                    onPremVideo.Title != o365Version.Title ||
                    onPremVideo.Description != o365Version.Description ||
                    onPremVideo.FileName != o365Version.FileName ||
                    onPremVideo.ServerRelativeUrl != o365Version.ServerRelativeUrl ||
                    onPremVideo.ThumbnailUrl != o365Version.ThumbnailUrl ||
                    onPremVideo.Url != o365Version.Url ||
                    onPremVideo.ViewCount != o365Version.ViewCount ||
                    onPremVideo.VideoDurationInSeconds != o365Version.VideoDurationInSeconds
                    )
                    {
                        onPremUtils.UpdateVideo(onPremVideo, o365Version, settings.ThumbnailWidth, settings.ThumbnailHeight, settings.ThumbnailHyperlinkFormat);
                    }
                }
            }
            foreach (Video o365Video in o365Videos)
            {
                var onPremVersion = onPremVideos.Find(opv => opv.ID == o365Video.ID);// find an entry oin the onprem list with the same videoid
                if (onPremVersion == null)
                {
                    onPremUtils.AddVideo(o365Video, settings.ThumbnailWidth, settings.ThumbnailHeight, settings.ThumbnailHyperlinkFormat);
                }
            }

            Console.WriteLine("****O365VideoSync started at " + DateTime.Now + " *****");
        }






    }



}
