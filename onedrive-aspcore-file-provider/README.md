# onedrive-aspcore-file-provider

This repository contains the OneDrive file system provider in ASP.NET Core for the File Manager component.

## Key Features

The following actions can be performed with OneDrive file system provider:

| **Actions** | **Description** |
| --- | --- |
| Read         | Reads the files and folders from OneDrive. |
| Details      | Gets the file's details like Type, Size, Location, and Modified date. |
| Download     | Downloads the selected file or folder from OneDrive. |
| Upload       | Uploads a file to OneDrive. It accepts uploaded media with the following characteristics: <ul><li>Maximum file size: 30MB</li><li>Accepted Media MIME types: `*/*`</li></ul> |
| Create       | Creates a new folder. |
| Delete       | Deletes a folder or file. |
| Copy         | Copies the selected files or folders to the target location. |
| Move         | Moves the selected files or folders to the desired location. |
| Rename       | Renames a folder or file. |
| Search       | Searches for files and folders using a text query. |

## Prerequisites

To run the service, configure the application to access OneDrive using the Microsoft Graph API. Register your application in [Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) and assign the required Microsoft Graph permissions, **Files.Read.All** and **Files.ReadWrite.All**, with admin consent.

Open the `appsettings.json` file and set your OneDrive account details under the `OneDriveSettings` section:

```json
"OneDriveSettings": {
  "TenantId": "<--Tenant Id-->",
  "ClientId": "<--Client Id-->",
  "ClientSecret": "<--Client Secret-->",
  "DriveId": "<--Drive Id (optional)-->"
}
```

If the `DriveId` is not specified, the provider gets the default drive associated with the authenticated user or application identity.

## How to run this application

To run this application, clone the `onedrive-aspcore-file-provider` repository and then navigate to its appropriate path where it has been located in your system.

To do so, open the command prompt and run the following commands one after the other.

```
git clone https://github.com/SyncfusionExamples/onedrive-aspcore-file-provider

cd onedrive-aspcore-file-provider
```

## Restore the NuGet package and build the application

To restore the NuGet package, run the following command in root folder of the application.

```
dotnet restore
```

To build the application, run the following command.

```
dotnet build
```

## Running application

After successful compilation, run the following command to run the application.

```
dotnet run
```

Now, the project will be hosted in http://localhost. To ensure the OneDrive-service, map the following URL in your browser.

```
http://localhost:<port-number>/api/test
```

## File Manager AjaxSettings

To access the basic actions such as Read, Delete, Copy, Move, Rename, Search, and Get Details of File Manager using OneDrive service, just map the following code snippet in the AjaxSettings property of File Manager.

Here, the `hostUrl` will be your locally hosted port number.

```
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDrive/OneDriveFileOperations'
}
```

## File download AjaxSettings

To perform download operation, initialize the `downloadUrl` property in ajaxSettings of the File Manager component.

```
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDrive/OneDriveFileOperations',
    downloadUrl: hostUrl + 'api/OneDrive/OneDriveDownload'
}
```

## File upload AjaxSettings

To perform upload operation, initialize the `uploadUrl` property in ajaxSettings of the File Manager component.

```
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDrive/OneDriveFileOperations',
    uploadUrl: hostUrl + 'api/OneDrive/OneDriveUpload'
}
```

## File image preview AjaxSettings

To perform image preview support in the File Manager component, initialize the `getImageUrl` property in ajaxSettings of the File Manager component.

```
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDrive/OneDriveFileOperations',
    getImageUrl: hostUrl + 'api/OneDrive/OneDriveGetImage'
}
```

The FileManager will be rendered as the following.

![File Manager](https://ej2.syncfusion.com/products/images/file-manager/readme.gif)

## Support

Product support is available through the following mediums:

* Creating incident in Syncfusion [Direct-trac](https://www.syncfusion.com/support/directtrac/incidents?utm_source=npm&utm_campaign=filemanager) support system or [Community forum](https://www.syncfusion.com/forums/essential-js2?utm_source=npm&utm_campaign=filemanager).

* New [GitHub issue](https://github.com/syncfusion/ej2-javascript-ui-controls/issues/new).

* Ask your query in [Stack Overflow](https://stackoverflow.com/?utm_source=npm&utm_campaign=filemanager) with tag `syncfusion` and `ej2`.

## License

Check the license details [here](https://github.com/syncfusion/ej2-javascript-ui-controls/blob/master/license).

## Changelog

Check the changelog [here](https://github.com/syncfusion/ej2-javascript-ui-controls/blob/master/controls/filemanager/CHANGELOG.md)

© Copyright 2020 Syncfusion, Inc. All Rights Reserved. The Syncfusion Essential Studio license and copyright applies to this distribution.
