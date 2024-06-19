using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;

namespace ModuleReplacement
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var clientContext = new ClientContext("http://w15-sp/sites/ftclab"))
            {
                XDocument settings = XDocument.Load("settings.xml");

                //Grab the master page gallery folder for uploading and the content types for the masterpage gallery 
                Web web = clientContext.Web;
                List gallery = web.GetCatalog(116);
                Folder folder = gallery.RootFolder;
                //Load info about the master page gallery
                clientContext.Load(folder);
                clientContext.Load(gallery,
                                    g => g.ContentTypes,
                                    g => g.RootFolder.ServerRelativeUrl);
                //Load the content types and master page info for the web
                clientContext.Load(web,
                                    w => w.ContentTypes,
                                    w => w.MasterUrl,
                                    w => w.CustomMasterUrl);
                clientContext.ExecuteQuery();
                //Get the list contentTypeId for the masterpage upload the master pages we know about
                const string parentMasterPageContentTypeId = "0x010105"; // Master Page 
                ContentType masterPageContentType = gallery.ContentTypes.FirstOrDefault(ct => ct.StringId.StartsWith(parentMasterPageContentTypeId));
                UploadAndSetMasterPages(web, folder, clientContext, settings, masterPageContentType.StringId);

                //Get the list contentTypeId for the PageLayouts upload the layout pages we know about
                const string parentPageLayoutContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811"; //Page Layout
                ContentType pageLayoutContentType = gallery.ContentTypes.FirstOrDefault(ct => ct.StringId.StartsWith(parentPageLayoutContentTypeId));
                UploadPageLayoutsAndUpdateReferences(web, folder, clientContext, settings, pageLayoutContentType.StringId);
            }
        }

        private static void UploadAndSetMasterPages(Web web, Folder folder, ClientContext clientContext, XDocument settings, string contentTypeId)
        {
            IList<MasterPageGalleryFile> masterPages = (from m in settings.Descendants("masterPage")
                                                        select new MasterPageGalleryFile
                                                        {
                                                            File = (string)m.Attribute("file"),
                                                            Replaces = (string)m.Attribute("replaces"),
                                                            ContentTypeId = contentTypeId
                                                        }).ToList();
            foreach (MasterPageGalleryFile masterPage in masterPages)
            {
                UploadAndSetMasterPage(web, folder, clientContext, masterPage);
            }
        }

        private static void UploadAndSetMasterPage(Web web, Folder folder, ClientContext clientContext, MasterPageGalleryFile masterPage)
        {
            using (var fileReadingStream = System.IO.File.OpenRead(masterPage.File))
            {
                //ensure that the masterpage is checked out if this is needed
                PublishingHelper.CheckOutFile(web, masterPage.File, folder.ServerRelativeUrl);

                //Use the FileCreationInformation to upload the new file
                var fileInfo = new FileCreationInformation();
                fileInfo.ContentStream = fileReadingStream;
                fileInfo.Overwrite = true;
                fileInfo.Url = masterPage.File;
                File file = folder.Files.Add(fileInfo);
                //Get the list item associated with the newly uploaded file
                ListItem item = file.ListItemAllFields;
                clientContext.Load(file.ListItemAllFields);
                clientContext.Load(file,
                                    f => f.CheckOutType,
                                    f => f.Level,
                                    f => f.ServerRelativeUrl);
                clientContext.ExecuteQuery();

                item["ContentTypeId"] = masterPage.ContentTypeId;
                item["UIVersion"] = Convert.ToString(15);
                item["MasterPageDescription"] = "Master Page Uploaded using CSOM";
                item.Update();
                clientContext.ExecuteQuery();
                //Check-in, publish and approve and the new masterpage file if needed
                PublishingHelper.CheckInPublishAndApproveFile(file);
                //Update the references to the replaced masterpage if needed
                PublishingHelper.SetMasterPageReferences(web, clientContext, masterPage, file);
            }
        }


        private static void UploadPageLayoutsAndUpdateReferences(Web web, Folder folder, ClientContext clientContext, XDocument settings, string contentTypeId)
        {
            //Load all the pageLayout replacements from the settings
            IList<LayoutFile> pageLayouts = (from m in settings.Descendants("pageLayout")
                                         select new LayoutFile
                                         {
                                             File = (string)m.Attribute("file"),
                                             Replaces = (string)m.Attribute("replaces"),
                                             Title = (string)m.Attribute("title"),
                                             ContentTypeId = contentTypeId,
                                             AssociatedContentTypeName = (string)m.Attribute("associatedContentTypeName"),
                                             DefaultLayout = m.Attribute("defaultLayout") != null && (bool)m.Attribute("defaultLayout")
                                         }).ToList();

            foreach (LayoutFile pageLayout in pageLayouts)
            {
                ContentType associatedContentType =
                    web.ContentTypes.FirstOrDefault(ct => ct.Name == pageLayout.AssociatedContentTypeName);
                pageLayout.AssociatedContentTypeId = associatedContentType.StringId;

                UploadPageLayout(web, folder, clientContext, pageLayout);
            }
            UpdatePages(web, clientContext, pageLayouts);
        }

        private static void UploadPageLayout(Web web, Folder folder, ClientContext clientContext, LayoutFile pageLayout)
        {
            using (var fileReadingStream = System.IO.File.OpenRead(pageLayout.File))
            {
                PublishingHelper.CheckOutFile(web, pageLayout.File, folder.ServerRelativeUrl);
                //Use the FileCreationInformation to upload the new file
                var fileInfo = new FileCreationInformation();
                fileInfo.ContentStream = fileReadingStream;
                fileInfo.Overwrite = true;
                fileInfo.Url = pageLayout.File;
                File file = folder.Files.Add(fileInfo);
                //Get the list item associated with the newly uploaded file
                ListItem item = file.ListItemAllFields;
                clientContext.Load(file.ListItemAllFields);
                clientContext.Load(file,
                                    f => f.CheckOutType,
                                    f => f.Level,
                                    f => f.ServerRelativeUrl);
                clientContext.ExecuteQuery();

                item["ContentTypeId"] = pageLayout.ContentTypeId;
                item["Title"] = pageLayout.Title;
                item["PublishingAssociatedContentType"] = string.Format(";#{0};#{1};#", pageLayout.AssociatedContentTypeName, pageLayout.AssociatedContentTypeId);
                item.Update();
                clientContext.ExecuteQuery();

                PublishingHelper.CheckInPublishAndApproveFile(file);

                PublishingHelper.UpdateAvailablePageLayouts(web, clientContext, pageLayout, file);

                if (pageLayout.DefaultLayout)
                {
                    PublishingHelper.SetDefaultPageLayout(web, clientContext, file);
                }
            }
        }

        private static void UpdatePages(Web web, ClientContext clientContext, IList<LayoutFile> pageLayouts)
        {
            //Get the Pages Library and all the list items it contains
            List pagesList = web.Lists.GetByTitle("Pages");
            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = pagesList.GetItems(allItemsQuery);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            foreach (ListItem item in items)
            {
                //Only update those pages that are using a page layout which is being replaced
                var pageLayout = item["PublishingPageLayout"] as FieldUrlValue;
                if (pageLayout != null)
                {
                    LayoutFile matchingLayout = pageLayouts.FirstOrDefault(p => pageLayout.Url.EndsWith("/" + p.Replaces));
                    if (matchingLayout != null)
                    {
                        //Check out the page so we can update the page layout being used
                        PublishingHelper.CheckOutFile(web, item);
                        //Update the pageLayout reference
                        pageLayout.Url = pageLayout.Url.Replace(matchingLayout.Replaces, matchingLayout.File);
                        item["PublishingPageLayout"] = pageLayout;
                        item.Update();
                        File file = item.File;
                        //Grab the file and other attributes so that we can check in etc.
                        clientContext.Load(file,
                            f => f.Level,
                            f => f.CheckOutType);
                        clientContext.ExecuteQuery();
                        //Check-in etc.
                        PublishingHelper.CheckInPublishAndApproveFile(file);
                    }
                }
            }
        }
    }
}
