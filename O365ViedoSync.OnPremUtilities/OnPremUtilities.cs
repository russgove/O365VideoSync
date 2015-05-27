using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using O365VideoSync.Common;
using System.Web;
using System.Net;
namespace O365VideoSync.OnPremUtilities
{
    public class OnPremUtilities
    {
        public string siteUrl { get; set; }
        public string listTitle { get; set; }
        public string userID { get; set; }
        public string password { get; set; }

        private ClientContext _clientContext;
        private ClientContext ClientContext
        {
            get
            {
                if (_clientContext != null)
                {
                    return _clientContext;
                }
                _clientContext = new ClientContext(siteUrl);
                _clientContext.Credentials = new NetworkCredential(userID, password);
                return _clientContext;
            }

        }


        public List<Video> GetVideos()
        {

            var web = ClientContext.Web;
            var list = web.Lists.GetByTitle(listTitle);
            var listItems = list.GetItems(new CamlQuery());
            ClientContext.Load(listItems);
            ClientContext.ExecuteQuery();
            var retrunVal = new List<Video>();
            foreach (ListItem item in listItems)
            {
                Video v = new Video();

                v.OnPremItemId = item.Id;
                v.ChannelID = (string)item["ChannelID"];
                v.CreatedDate = (DateTime)item["CreatedDate"];
                v.Description = (string)item["Description"];
                v.DisplayFormUrl = (string)item["DisplayFormUrl"];
                v.FileName = (string)item["FileName"];
                v.OwnerName = (string)item["OwnerName"];
                v.ServerRelativeUrl = (string)item["ServerRelativeUrl"];
                v.ThumbnailUrl = (string)item["ThumbnailUrl"];
                v.Title = (string)item["Title"];
                v.ID = (string)item["O365VideoID"];
                v.Url = (string)item["Url"];
                v.VideoDurationInSeconds = (item["VideoDurationInSeconds"] == null) ? 0 : (double)item["VideoDurationInSeconds"];
                v.VideoProcessingStatus = (item["VideoProcessingStatus"] == null) ? 0 : (double)item["VideoProcessingStatus"];
                v.ViewCount = (item["ViewCount"] == null) ? 0 : (double)item["ViewCount"];
                v.YammerObjectUrl = (string)item["YammerObjectUrl"];
                retrunVal.Add(v);
            }

            return retrunVal;


        }

        public void RemoveVideo(Video onPremVideo)
        {
            var web = ClientContext.Web;
            var list = web.Lists.GetByTitle(listTitle);
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = list.GetItemById(onPremVideo.OnPremItemId);
            oListItem.Recycle();
            ClientContext.ExecuteQuery();
        }

        public void UpdateVideo(Video onPrenVersion, Video o365Version, int thumbnailWidth, int thumnnailHeight, string thumbnailHyperlinkFormat)
        {
            var web = ClientContext.Web;
            var list = web.Lists.GetByTitle(listTitle);
            ListItem oListItem = list.GetItemById(onPrenVersion.OnPremItemId);
            ClientContext.Load(oListItem           );
    //        ClientContext.Load(oListItem,
    //item => item["Title"],
    //item => item["ThumbnailUrl"]
    //);

            ClientContext.ExecuteQuery();
            MoveVideoFieldsToListItem(o365Version, oListItem, thumbnailWidth, thumnnailHeight, thumbnailHyperlinkFormat);
            oListItem.Update();
            ClientContext.ExecuteQuery();
        }


        public void AddVideo(Video o365Video, int thumbnailWidth, int thumnnailHeight, string thumbnailHyperlinkFormat)
        {
            var web = ClientContext.Web;
            var list = web.Lists.GetByTitle(listTitle);
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = list.AddItem(itemCreateInfo);
            MoveVideoFieldsToListItem(o365Video, oListItem, thumbnailWidth, thumnnailHeight, thumbnailHyperlinkFormat);
            oListItem.Update();
            ClientContext.ExecuteQuery();

        }
        private static void MoveVideoFieldsToListItem(Video o365Version, ListItem oListItem, int thumbnailWidth, int thumnnailHeight, string thumbnailHyperlinkFormat)
        {
            oListItem["Title"] = o365Version.Title;
            oListItem["Description"] = o365Version.Description;
            oListItem["ChannelID"] = o365Version.ChannelID;
            oListItem["CreatedDate"] = o365Version.CreatedDate;
            oListItem["DisplayFormUrl"] = o365Version.DisplayFormUrl;
            oListItem["FileName"] = o365Version.FileName;
            oListItem["OwnerName"] = o365Version.OwnerName;
            oListItem["ServerRelativeUrl"] = o365Version.ServerRelativeUrl;
            oListItem["ThumbnailUrl"] = o365Version.ThumbnailUrl;
            oListItem["Url"] = o365Version.Url;
            oListItem["VideoDurationInSeconds"] = o365Version.VideoDurationInSeconds;
            oListItem["VideoProcessingStatus"] = o365Version.VideoProcessingStatus;
            oListItem["ViewCount"] = o365Version.ViewCount;
            oListItem["O365VideoID"] = o365Version.ID;
            oListItem["Thumbnail"] = new FieldUrlValue() { Url = String.Format("{0}&width={1}&height={2}", o365Version.ThumbnailUrl, thumbnailWidth, thumnnailHeight), Description = o365Version.Title };
            oListItem["LinkToVideo"] = new FieldUrlValue() { Url = o365Version.DisplayFormUrl, Description = o365Version.Title };
            oListItem["ThumbnailAsHyperlink"] = String.Format(thumbnailHyperlinkFormat, o365Version.YammerObjectUrl, o365Version.ThumbnailUrl, thumbnailWidth, thumnnailHeight, o365Version.Title);

        }

        public bool ListExists(string listTitle)
        {
            var web = ClientContext.Web;
            ListCollection listCollection = ClientContext.Web.Lists;
            ClientContext.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            ClientContext.ExecuteQuery();
            return (listCollection.Count > 0);
        }

        public void CreateList(string listTitle, string listDescription, bool onQuickLaunch)
        {
            var web = ClientContext.Web;
            ListCreationInformation listCreationInfo = new ListCreationInformation();
            listCreationInfo.Title = listTitle;
            listCreationInfo.Description = listDescription;
            listCreationInfo.QuickLaunchOption = (onQuickLaunch) ? QuickLaunchOptions.On : QuickLaunchOptions.Off;

            listCreationInfo.TemplateType = (int)ListTemplateType.GenericList;
            List list = web.Lists.Add(listCreationInfo);
            ClientContext.ExecuteQuery();
            Field ChannelID = AddField(list, "Channel ID", "ChannelID", "Text", false);
            Field O365VideoID = AddField(list, "Video ID", "O365VideoID", "Text", false);
            Field CreatedDate = AddField(list, "Created Date", "CreatedDate", "DateTime", false);
            Field Description = AddField(list, "Description", "Description", "Note", true);
            Field DisplayFormUrl = AddField(list, "Display Form Url", "DisplayFormUrl", "Note", true);
            Field FileName = AddField(list, "File Name", "FileName", "Text", false);
            Field ServerRelativeUrl = AddField(list, "Server Relative Url", "ServerRelativeUrl", "Text", false);
            Field ThumbnailUrl = AddField(list, "Thumbnail Url", "ThumbnailUrl", "Text", false);
            Field Url = AddField(list, "Url", "Url", "Text", false);



            Field VideoDurationInSeconds = AddField(list, "Video Duration In Seconds", "VideoDurationInSeconds", "Number", false);
            Field VideoProcessingStatus = AddField(list, "Processing Status", "VideoProcessingStatus", "Number", true);
            Field VideoCount = AddField(list, "Views", "ViewCount", "Number", false);

            Field LinkToVideo = AddField(list, "Link To Video", "LinkToVideo", "URL", false);
            Field Thumbnail = AddField(list, "Thumbnail", "Thumbnail", "URL", false, "Image");




            Field YammerObjectUrl = AddField(list, "Yammer Object Url", "YammerObjectUrl", "Text", false);
            Field OwnerName = AddField(list, "Owner Name", "OwnerName", "Note", false);
            Field ThumbnailAsHyperlink = AddField(list, "ThumbnailAsHyperlink", "ThumbnailAsHyperlink", "Note", true, "", true, "FullHtml", true);


            list.Update();
            ClientContext.Load(list);
            var x = list.Fields;
            ClientContext.Load(x);

            ClientContext.ExecuteQuery();



            list.Views.Add(new ViewCreationInformation
            {
                Title = "ThumbnailView",
                ViewTypeKind = ViewType.Html,
                ViewFields = new string[] { "Edit", "LinkToVideo", "ThumbnailAsHyperlink", "Description" },
                SetAsDefaultView = true,
                RowLimit = 25,
                PersonalView = false,
                Paged = true,
            });
            list.Update();
            ClientContext.ExecuteQuery();


        }

        private Field AddField(List list, string displayName, string name, string type, bool includeInDefaultView, string format = "", bool richText = false, string richTextType = "", bool isoldateStyles = false)
        {


            string FieldXMLFormat = "<Field DisplayName='{0}' StaticName='{1}' Name='{2}' Type='{3}' Format='{4}' RichText='{5}' RichTextMode='{6}' IsolateStyles='{7}' />";
            string FieldXML = string.Format(FieldXMLFormat, displayName, name, name, type, format, (richText) ? "TRUE" : "FALSE", richTextType, (isoldateStyles) ? "TRUE" : "FALSE");
            Field fld = list.Fields.AddFieldAsXml(FieldXML, includeInDefaultView, AddFieldOptions.AddFieldInternalNameHint);
            return fld;

        }
    }
}
