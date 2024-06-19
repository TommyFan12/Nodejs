using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;

namespace Core.ListTemplateReplacement
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var clientContext = new ClientContext("http://w15-sp/sites/ftclab"))
            {
                Web web = clientContext.Web;
                ListCollection listCollection = web.Lists;
                clientContext.Load(listCollection,
                                    l => l.Include(list => list.BaseTemplate,
                                                    list => list.BaseType,
                                                    list => list.Title));
                clientContext.ExecuteQuery();
                var listsToReplace = new List<List>();
                foreach (List list in listCollection)
                {
                    //10003 is the Template Id for the custom list template instances we're replacing
                    if (list.BaseTemplate == 10003)
                    {
                        listsToReplace.Add(list);
                    }
                }
                foreach (List list in listsToReplace)
                {
                    ReplaceList(clientContext, listCollection, list);
                }
            }
        }

        private static void ReplaceList(ClientContext clientContext, ListCollection listCollection, List listToBeReplaced)
        {
            //This is to let me re-run a bunch of times, this needs to come out before we handover
            try
            {
                List deleteMe = listCollection.GetByTitle("ContosoLibraryApp");
                deleteMe.DeleteObject();
                clientContext.ExecuteQuery();
            }
            catch (Exception)
            {
                //swallow
            }

            var newList = CreateReplacementList(clientContext, listCollection, listToBeReplaced);

            SetListSettings(clientContext, listToBeReplaced, newList);

            SetContentTypes(clientContext, listToBeReplaced, newList);

            AddViews(clientContext, listToBeReplaced, newList);

            RemoveViews(clientContext, listToBeReplaced, newList);

            MigrateContent(clientContext, listToBeReplaced, newList);
        }

        private static List CreateReplacementList(ClientContext clientContext, ListCollection lists, List listToBeReplaced)
        {
            var creationInformation = new ListCreationInformation
            {
                Title = listToBeReplaced.Title + "App",
                TemplateType = (int)ListTemplateType.DocumentLibrary,
            };
            List newList = lists.Add(creationInformation);
            clientContext.ExecuteQuery();
            return newList;
        }

        private static void SetListSettings(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            clientContext.Load(listToBeReplaced, 
                                l => l.EnableVersioning, 
                                l => l.EnableModeration, 
                                l => l.EnableMinorVersions,
                                l => l.DraftVersionVisibility );
            clientContext.ExecuteQuery();
            newList.EnableVersioning = listToBeReplaced.EnableVersioning;
            newList.EnableModeration = listToBeReplaced.EnableModeration;
            newList.EnableMinorVersions= listToBeReplaced.EnableMinorVersions;
            newList.DraftVersionVisibility = listToBeReplaced.DraftVersionVisibility;
            newList.Update();
            clientContext.ExecuteQuery();
        }


        private static void SetContentTypes(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            clientContext.Load(listToBeReplaced,
                                l => l.ContentTypesEnabled,
                                l => l.ContentTypes);
            clientContext.Load(newList,
                                l => l.ContentTypesEnabled,
                                l => l.ContentTypes);
            clientContext.ExecuteQuery();

            //If the originat list doesn't use ContentTypes there's nothing to do here.
            if (!listToBeReplaced.ContentTypesEnabled) return;

            newList.ContentTypesEnabled = true;
            newList.Update();
            clientContext.ExecuteQuery();
            foreach (var contentType in listToBeReplaced.ContentTypes)
            {
                if (!newList.ContentTypes.Any(ct => ct.Name == contentType.Name))
                {
                    //current Content Type needs to be added to new list
                    //Note that the Parent is used as contentType is the list instance not the site instance.
                    newList.ContentTypes.AddExistingContentType(contentType.Parent);
                    newList.Update();
                    clientContext.ExecuteQuery();
                }
            }
            //We need to re-load the ContentTypes for newList as they may have changed due to an add call above
            clientContext.Load(newList, l => l.ContentTypes);
            clientContext.ExecuteQuery();
            //Remove any content type that are not needed
            var contentTypesToDelete = new List<ContentType>();
            foreach (var contentType in newList.ContentTypes)
            {
                if (!listToBeReplaced.ContentTypes.Any(ct => ct.Name == contentType.Name))
                {
                    //current Content Type needs to be removed from new list
                    contentTypesToDelete.Add(contentType);
                }
            }
            foreach (var contentType in contentTypesToDelete)
            {
                contentType.DeleteObject();
            }
            newList.Update();
            clientContext.ExecuteQuery();
        }

        private static void AddViews(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            ViewCollection views = listToBeReplaced.Views;
            clientContext.Load(views,
                                v => v.Include(view => view.Paged,
                                    view => view.PersonalView,
                                    view => view.ViewQuery,
                                    view => view.Title,
                                    view => view.RowLimit,
                                    view => view.DefaultView,
                                    view => view.ViewFields,
                                    view => view.ViewType));
            clientContext.Load(newList.Views, 
                                v => v.Include(view => view.Title));
            clientContext.ExecuteQuery();

            //Build a list of views which only exist on the source list
            var viewsToCreate = new List<ViewCreationInformation>();
            foreach (View view in listToBeReplaced.Views)
            {
                if (!newList.Views.Any(v => v.Title == view.Title))
                {
                    var createInfo = new ViewCreationInformation
                    {
                        Paged = view.Paged,
                        PersonalView = view.PersonalView,
                        Query = view.ViewQuery,
                        Title = view.Title,
                        RowLimit = view.RowLimit,
                        SetAsDefaultView = view.DefaultView,
                        ViewFields = view.ViewFields.ToArray(),
                        ViewTypeKind = GetViewType(view.ViewType),
                    };
                    viewsToCreate.Add(createInfo);
                }
            }
            //Create the list that we need to
            foreach (ViewCreationInformation newView in viewsToCreate)
            {
                newList.Views.Add(newView);
            }
            newList.Update();
        }

        private static void RemoveViews(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            //Update the list of views
            clientContext.Load(newList, l => l.Views);
            clientContext.ExecuteQuery();

            var viewsToRemove = new List<View>();
            foreach (View view in newList.Views)
            {
                if (!listToBeReplaced.Views.Any(v => v.Title == view.Title))
                {
                    //new list contains a view which is not on the source list, remove it
                    viewsToRemove.Add(view);
                }
            }
            foreach (View view in viewsToRemove)
            {
                view.DeleteObject();
            }
            newList.Update();
            clientContext.ExecuteQuery();
        }

        private static void MigrateContent(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            ListItemCollection items = listToBeReplaced.GetItems(CamlQuery.CreateAllItemsQuery());
            Folder destination = newList.RootFolder;
            Folder source = listToBeReplaced.RootFolder;
            clientContext.Load(destination,
                                d => d.ServerRelativeUrl);
            clientContext.Load(source,
                                s => s.Files,
                                s => s.ServerRelativeUrl);
            clientContext.Load(items,
                                i => i.IncludeWithDefaultProperties(item => item.File));
            clientContext.ExecuteQuery();


            foreach (File file in source.Files)
            {
                string newUrl = file.ServerRelativeUrl.Replace(source.ServerRelativeUrl, destination.ServerRelativeUrl);
                file.CopyTo(newUrl, true);
//                file.MoveTo(newUrl, MoveOperations.Overwrite);
            }
            clientContext.ExecuteQuery();
        }

        private static ViewType GetViewType(string viewType)
        {
            switch (viewType)
            {
                case "HTML":
                    return ViewType.Html;
                case "GRID":
                    return ViewType.Grid;
                case "CALENDAR":
                    return ViewType.Calendar;
                case "RECURRENCE":
                    return ViewType.Recurrence;
                case "CHART":
                    return ViewType.Chart;
                case "GANTT":
                    return ViewType.Gantt;
                default:
                    return ViewType.None;
            }
        }
    }
}
