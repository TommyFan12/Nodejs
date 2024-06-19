using System;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;

namespace ModuleReplacement
{
    public class PublishingHelper
    {
        public static void CheckInPublishAndApproveFile(File uploadFile)
        {
            if (uploadFile.CheckOutType != CheckOutType.None)
            {
                uploadFile.CheckIn("Updating branding", CheckinType.MajorCheckIn);
            }

            if (uploadFile.Level == FileLevel.Draft)
            {
                uploadFile.Publish("Updating branding");
            }

            uploadFile.Context.Load(uploadFile, f => f.ListItemAllFields);
            uploadFile.Context.ExecuteQuery();

            if (uploadFile.ListItemAllFields["_ModerationStatus"].ToString() == "2") // SPModerationStatusType.Pending
            {
                uploadFile.Approve("Updating branding");
                uploadFile.Context.ExecuteQuery();
            }
        }

        public static void CheckOutFile(Web web, string fileName, string filePath)
        {
            var fileUrl = String.Concat(filePath, "/", fileName);
            var temp = web.GetFileByServerRelativeUrl(fileUrl);

            CheckOutFile(web, temp);
        }

        public static void CheckOutFile(Web web, ListItem item)
        {
            var file = item.File;
            CheckOutFile(web, file);
        }

        private static void CheckOutFile(Web web, File temp)
        {
            web.Context.Load(temp, f => f.Exists);
            web.Context.ExecuteQuery();

            if (temp.Exists)
            {
                web.Context.Load(temp, f => f.CheckOutType);
                web.Context.ExecuteQuery();

                if (temp.CheckOutType != CheckOutType.None)
                {
                    temp.UndoCheckOut();
                }

                temp.CheckOut();
                web.Context.ExecuteQuery();
            }
        }

        public static void SetMasterPageReferences(Web web, ClientContext clientContext, MasterPageGalleryFile masterPageGalleryFile, File file)
        {
            //Make apropriate master page replacements
            if (web.MasterUrl.EndsWith("/" + masterPageGalleryFile.Replaces))
            {
                web.MasterUrl = file.ServerRelativeUrl;
            }
            if (web.CustomMasterUrl.EndsWith("/" + masterPageGalleryFile.Replaces))
            {
                web.CustomMasterUrl = file.ServerRelativeUrl;
            }
            web.Update();
            clientContext.ExecuteQuery();
        }

        public static void UpdateAvailablePageLayouts(Web web, ClientContext clientContext, LayoutFile pageLayout, File newLayout)
        {
            PropertyValues webProperties = web.AllProperties;
            clientContext.Load(webProperties);
            clientContext.ExecuteQuery();

            string pageLayouts = (string)webProperties["__PageLayouts"];
            XDocument document = XDocument.Parse(pageLayouts);
            XElement layoutToUpdate = document.Descendants("layout").FirstOrDefault(el => el.Attribute("url").Value.EndsWith("/" + pageLayout.Replaces));
            if (layoutToUpdate != null)
            {
                layoutToUpdate.Attribute("guid").SetValue(newLayout.ListItemAllFields["UniqueId"]);
                layoutToUpdate.Attribute("url").SetValue(newLayout.ServerRelativeUrl);
            }
            webProperties["__PageLayouts"] = document.ToString();
            web.Update();
            clientContext.ExecuteQuery();

        }

        public static void SetDefaultPageLayout(Web web, ClientContext clientContext, File pageLayoutFile)
        {
            PropertyValues webProperties = web.AllProperties;
            clientContext.Load(webProperties);
            clientContext.ExecuteQuery();

            string guid = pageLayoutFile.ListItemAllFields["UniqueId"].ToString();
            string url = pageLayoutFile.ServerRelativeUrl;
            webProperties["__DefaultPageLayout"] = String.Format("<layout guid=\"{0}\" url=\"{1}\" />", guid, url);
            web.Update();
            clientContext.ExecuteQuery();

        }
    }
}