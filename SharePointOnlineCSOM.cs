using System;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Collections.Generic;

//Requires NuGet pkg https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

namespace SharePointOnlineCSOMDemo
{
    public class SharePointOnlineCSOM
    {

        /*
         * Parameters:
         * 'sharePointSiteUrl' e.g. "https://example.sharepoint.com/examplesite". Note, site must already exist
         * 'userName' user with site admin
         * 'password' user with site admin
         * 
         * Returns:
         * client context object
        */
        public ClientContext ConnectToSharePoint(string sharePointSiteUrl, string userName, string password)
        {
            ClientContext context = new ClientContext(sharePointSiteUrl);
            context.AuthenticationMode = ClientAuthenticationMode.Default;
            context.Credentials = new SharePointOnlineCredentials(userName, GetSPOSecureStringPassword(password));
            return context;
        }

        /*
         * Generate secure password from string password
        */
        private SecureString GetSPOSecureStringPassword(string password)
        {
            var secureString = new SecureString();
            foreach (char c in password)
            {
                secureString.AppendChar(c);
            }
            return secureString;
        }

        /*
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library to be created
        */
        public void CreateDocumentLibrary(ClientContext context, string documentLibraryName)
        {
            //Probably should check if Document Library already exists rather than just creating it
            try
            {
                Web web = context.Web;
                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = documentLibraryName;
                creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
                List list = web.Lists.Add(creationInfo);
                list.Description = $"{documentLibraryName}";
                list.Update();
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                //Document Library probably already exists. 
            }
        }

        /*
         * Adds a text column to a document library
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library (should already exist)
         * 'textColumnName' name of the text column to create in the document library
        */
        public void AddTextColumnToDocumentLibrary(ClientContext context, string documentLibraryName, string textColumnName)
        {
            bool columnExists = false;
            Web web = context.Web;
            var list = web.Lists.GetByTitle(documentLibraryName);
            context.Load(list.Fields);
            context.ExecuteQuery();

            foreach (Field f in list.Fields)
            {
                if (f.Title == textColumnName) //Found a match
                {
                    columnExists = true;
                    break; //Don't bother looking any further; exit out of inner foreach loop
                }
            }

            //If column does not exist in Document Library, create it.
            if (!columnExists)
            {
                // Create a column of Type='Text'                                       
                list.Fields.AddFieldAsXml($"<Field ID='{Guid.NewGuid()}' Type='Text' DisplayName='{textColumnName}' Name='{textColumnName}'/>", true, AddFieldOptions.AddFieldInternalNameHint);
                list.Update();
                context.ExecuteQuery();
            }
        }

        /*
         * Adds a lookup column to a document library and links to an SharePoint List
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library (should already exist)
         * 'srcListName' name of the list that will be connected to the lookup column (should already exist)
         * 'lookupColumnName' name of the text column to create in the document library
        */
        public void AddLookupColumnToDocumentLibrary(ClientContext context, string documentLibraryName, string srcListName, string lookupColumnName)
        {
            try
            {
                //Lookup column internal name using DisplayName
                bool columnExists = false;
                Web web = context.Web;
                var list = web.Lists.GetByTitle(documentLibraryName);
                context.Load(list.Fields);
                context.ExecuteQuery();

                foreach (Field f in list.Fields)
                {
                    if (f.Title == lookupColumnName) //Found a match
                    {
                        columnExists = true;
                        break; //Don't bother looking any further
                    }
                }

                //If column does not exist in Document Library, create it.
                if (!columnExists)
                {
                    List targetList = context.Web.Lists.GetByTitle(srcListName);
                    context.Load(targetList, l => l.Id);
                    context.ExecuteQuery();

                    string addColumnXml = $"<Field ID='{Guid.NewGuid()}' Type='Lookup' DisplayName='{lookupColumnName}' Name='{lookupColumnName}' StaticName='{lookupColumnName}' List='{targetList.Id}' ShowField='Title'/>";
                    list.Fields.AddFieldAsXml(addColumnXml, true, AddFieldOptions.AddFieldInternalNameHint);
                    list.Update();
                    context.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                //Handle exception
            }
        }

        /*
         * Copies a file between two document libraries in the same SharePoint site
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'srcDocumentLibraryName' name of the document library to where the file is located
         * 'destDocumentLibraryName' name of the document library where the file will be copied to
         * 'fileName' name of the file to be copied
        */
        public void CopyDocument(ClientContext context, string srcDocumentLibraryName, string destDocumentLibraryName, string fileName)
        {
            Web web = context.Web;
            List srcList = web.Lists.GetByTitle(srcDocumentLibraryName);
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>{fileName}</Value></Eq></Where></Query></View>";
            ListItemCollection itemColl = srcList.GetItems(camlQuery);
            context.Load(itemColl);
            context.ExecuteQuery();

            //Get the destination list
            Web destWeb = context.Web;
            context.Load(destWeb);
            context.ExecuteQuery();

            foreach (var doc in itemColl)
            {
                try
                {
                    //Get the file
                    Microsoft.SharePoint.Client.File file = doc.File;
                    context.Load(file);
                    context.ExecuteQuery();

                    //Assemble relative path
                    string relativePath = destWeb.ServerRelativeUrl.TrimEnd('/') + "/" + destDocumentLibraryName + "/" + file.Name;

                    //Upload file to target document library
                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, relativePath, fileInfo.Stream, true);
                }
                catch (Exception ex)
                {
                    //Handle exception
                }
            }
        }

        /*
         * Uploads a local file to a document library. Supports large files >2MB
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library to where the file will be uploaded to
         * 'filePath' the path to the local file including the file name (e.g. "c:\\temp\\test.doc")
         * 'overwrite' set to true if you want to overwrite file if already exists; otherwise false
        */
        public void UploadFileToSharePoint(ClientContext context, string documentLibraryName, string filePath, bool overwrite)
        {
            try
            {
                //Extract file name from filePath
                int length2 = filePath.Length;
                int lastSlash = filePath.LastIndexOf("\\");
                lastSlash = lastSlash + 1;
                int relativeLength2 = length2 - lastSlash;
                string fileName = filePath.Substring(lastSlash, relativeLength2);

                //Get destination
                Web destWeb = context.Web;
                context.Load(destWeb);
                context.ExecuteQuery();

                //Assemble relative path
                string relativePath = destWeb.ServerRelativeUrl.TrimEnd('/') + "/" + documentLibraryName + "/" + fileName;

                //Upload file to target document library
                using (var fileStream = new FileStream(filePath, FileMode.Open))
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, relativePath, fileStream, overwrite); 
                }
            }
            catch (Exception ex)
            {
                //Handle exception
            }
        }

        /*
         * Gets item Id for a document in a document library
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library 
         * 'fileName' for which the itemId will be returned 
         * 
         * Returns:
         * ItemId (int)
        */
        public int GetItemId(ClientContext context, string documentLibraryName, string fileName)
        {
            int itemId = 0;

            Web web = context.Web;
            List targetList = web.Lists.GetByTitle(documentLibraryName);
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>{fileName}</Value></Eq></Where></Query></View>";
            ListItemCollection listItems = targetList.GetItems(camlQuery);
            context.Load(listItems);
            context.ExecuteQuery();

            if (listItems.Count == 0 || listItems.Count > 2)
            {
                Console.WriteLine("Ambiguous results returned from CAML query.");
            }
            else
            {
                foreach (ListItem item in listItems)
                {
                    //Should only return one row
                    itemId = item.Id;
                }
            }

            return itemId;
        }

        /*
         * Applies a value (aka metadata) to a text column for a given item in a document library
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library 
         * 'columnName' is the name of the column 
         * 'value' is the string value to apply in the column for the itemId
         * 'itemId' is the item id of the document (row) in the document library. Use GetItemId() to find this.
        */
        public void ApplyTextColumnMetadataToSharePointFile(ClientContext context, string documentLibraryName, string columnName, string value, int itemId)
        {
            try
            {
                Web web = context.Web;
                var list = web.Lists.GetByTitle(documentLibraryName);
                context.Load(list.Fields);
                context.ExecuteQuery();

                foreach (Field f in list.Fields)
                {
                    if (f.Title == columnName) //Found a match
                    {
                        string internalColumnName = f.InternalName;
                        Microsoft.SharePoint.Client.List oList = context.Web.Lists.GetByTitle(documentLibraryName);
                        ListItem oListItem = oList.GetItemById(itemId);
                        oListItem[internalColumnName] = value;
                        oListItem.Update();
                        context.ExecuteQuery();
                    }
                    else
                    {
                        //No match found
                    }
                }
            }
            catch (Exception ex)
            {
                //Handle exception
            }
        }

        /*
         * Applies a value (aka metadata) to a text column for a given item in a document library
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'documentLibraryName' name of the document library 
         * 'columnName' is the name of the column 
         * 'itemId' is the item id of the document (row) in the document library
         * 'listItemId' is the item id from the lookup list. Use an Id returned from GetLookupListItems() Dictionary.
        */
        public void ApplyLookupColumnMetadataToSharePointFile(ClientContext context, string documentLibraryName, string columnName, int itemId, int listItemId)
        {
            try
            {
                //Lookup column internal name using DisplayName
                Web web = context.Web;
                var list = web.Lists.GetByTitle(documentLibraryName);
                context.Load(list.Fields);
                context.ExecuteQuery();

                foreach (Field f in list.Fields)
                {
                    if (f.Title == columnName) //Found a match
                    {
                        string internalColumnName = f.InternalName;

                        Microsoft.SharePoint.Client.List oList = context.Web.Lists.GetByTitle(documentLibraryName);
                        ListItem oListItem = oList.GetItemById(itemId);

                        //Credit to http://sharepoint.stackexchange.com/questions/8017/how-to-set-listitem-lookup-field-value
                        oListItem[internalColumnName] = listItemId;
                        oListItem.Update();
                        context.ExecuteQuery();
                    }
                    else
                    {
                        //No match found
                    }
                }
            }
            catch (Exception ex)
            {
                //Handle error
            }
        }

        /*
         * Gets the ID and Name of each item in a SharePoint list.
         * Parameters:
         * 'context' client context object returned from ConnectToSharePoint()
         * 'listName' name of the SharePoint list 
         * 
         * Returns:
         * IDictionary<int, string>
        */
        public IDictionary<int, string> GetLookupListItems(ClientContext context, string listName)
        {
            Dictionary<int, string> lookupList = new Dictionary<int, string>();

            try
            {
                Web web = context.Web;
                List list = web.Lists.GetByTitle(listName);
                ListItemCollection listItemCollection = list.GetItems(CamlQuery.CreateAllItemsQuery());

                context.Load(listItemCollection,
                           eachItem => eachItem.Include(
                            item => item,
                            item => item["ID"],
                            item => item["Title"]));
                context.ExecuteQuery();

                foreach (ListItem listItem in listItemCollection)
                {
                    int Id = System.Int32.Parse(listItem["ID"].ToString());
                    lookupList.Add(Id, listItem["Title"].ToString());
                }
            }
            catch (Exception ex)
            {
                //Handle error
            }
            return lookupList;
        }
    }
}
