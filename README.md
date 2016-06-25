# SharePoint Online CSOM Class
A C# class to demonstrate working with SharePoint Online using Client Side Object Model (CSOM).

Note: This class has little/no error handling and is for demonstration purposes only. There are no warranties, expressed or implied.

# Examples

**//Authenticating to SharePoint Online.**

`SharePointOnlineCSOM spObject = new SharePointOnlineCSOM();`

`ClientContext context = spObject.ConnectToSharePoint("https://examplesite.sharepoint.com/demosite", "user@domain.com", "some_password");`

**//Creating a Document Library**

`spObject.CreateDocumentLibrary(context, "My Document Library");`

**//Adding a text column to a Document Library**

`spObject.AddTextColumnToDocumentLibrary(context, "My Document Library", "Location");`

**//Adding a lookup column to a Document Library that is connected to a list called "Days of the Week"**

`spObject.AddLookupColumnToDocumentLibrary(context, "My Document Library", "Days of the Week", "Day");`

**//Copy a file from one Document Library to another Document Library**

`spObject.CopyDocument(context, "My Document Library", "My Document Library 2", "test.doc");`

**//Upload a local file to a SharePoint Document Library**

`spObject.UploadFileToSharePoint(context, "My Document Library", "c:\\temp\\test.csv", true);`

**//Get an itemId for a Document in a Document Library and then apply a value to it on a column**

`int itemId = spObject.GetItemId(context, "My Document Library", "test2.json");`
`spObject.ApplyTextColumnMetadataToSharePointFile(context, "My Document Library", "Test Column", "Test value", itemId);`

**//Get the items in a SharePoint lookup list, select an item from this list and apply the value to the lookup column in a Document Library**

`int listItemId = 0;`

`string lookupListName = "Days of the Week";`

`IDictionary<int, string> dict = spObject.GetLookupListItems(context, lookupListName);`

`foreach (KeyValuePair<int, string> entry in dict)`

`{`

    `if (entry.Value.ToString() == "Tuesday")`
    
    `{`
        
        `listItemId = entry.Key;`
        
    `}`
    
`}`

`int itemId = spObject.GetItemId(context, "My Document Library", "test.doc");`

`spObject.ApplyLookupColumnMetadataToSharePointFile(context, "My Document Library", "Day", itemId, listItemId);`

