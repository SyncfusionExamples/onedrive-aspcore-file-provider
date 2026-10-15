# onedrive-aspcore-file-provider

This repository contains the **OneDrive file system provider** in ASP.NET Core for the Syncfusion [File Manager component](https://www.syncfusion.com/ej2-asp-core-file-manager). It enables the File Manager to browse, manage, and perform file operations directly against a Microsoft OneDrive drive through the [Microsoft Graph API](https://developer.microsoft.com/en-us/graph).

## Key Features

The following actions can be performed using the OneDrive file system provider:

| **Actions** | **Description** |
| --- | --- |
| **Read** | Reads the files and folders from OneDrive, including navigation breadcrumbs and item metadata. |
| **Details** | Gets a file or folder's details such as Type, Size, Location, and Modified date. |
| **Download** | Downloads the selected file or folder as a ZIP archive from OneDrive. |
| **Upload** | Uploads one or more files to OneDrive. It supports drag-and-drop, multiple files, and large request sizes. |
| **Create** | Creates a new folder at the selected location. |
| **Delete** | Deletes the selected folders or files. |
| **Copy** | Copies the selected files or folders to a target location, with conflict handling. |
| **Move** | Moves the selected files or folders to the desired location, with conflict handling. |
| **Rename** | Renames a folder or file. |
| **Search** | Searches for files and folders using a text query within the current context. |

## Prerequisites

This provider uses **app-only (client credentials) authentication**. It must be registered as an Azure Active Directory application and granted the following Microsoft Graph **application permissions** with admin consent:

- **Files.Read.All**
- **Files.ReadWrite.All**

For details on registering an application, see [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).

> **Important:** The Azure AD application and the OneDrive storage account must belong to the **same directory (tenant)**. The provider authenticates as the application principal and accesses the drive under that tenant.

> **Note:** Because the provider authenticates with application permissions (no signed-in user), the `/me` endpoint cannot be used for drive discovery. A valid **DriveId** must therefore be configured explicitly in `appsettings.json`. The application validates the required settings at startup and fails fast with a clear configuration error if `TenantId`, `ClientId`, `ClientSecret`, or `DriveId` is missing.

## Configuration

Open `appsettings.json` and set your OneDrive account details under the `OneDriveSettings` section:

```json
{
  "OneDriveSettings": {
    "TenantId": "<--Tenant Id-->",
    "ClientId": "<--Client Id-->",
    "ClientSecret": "<--Client Secret-->",
    "DriveId": "<--Drive Id-->",
    "RootFolderId": "<--Root Folder Id-->"
  }
}
```

| Setting | Description |
|---|---|
| `TenantId` | Azure AD tenant ID of the registered application. |
| `ClientId` | Application (client) ID of the registered application. |
| `ClientSecret` | Client secret value created for the registered application. |
| `DriveId` | **Required.** The OneDrive drive ID to access. Because the provider uses app-only authentication, the drive cannot be resolved from a signed-in user and must be specified explicitly. |
| `RootFolderId` | Optional. The OneDrive item ID of a specific folder to scope the File Manager root. When set, the File Manager will display only this folder and its contents. |

## Nested Root Folder Support

The provider supports loading the File Manager starting from a specific OneDrive folder instead of the drive root. To enable this, set the `RootFolderId` value in `appsettings.json` with the OneDrive item ID of the folder you want to use as the root.

When `RootFolderId` is configured:

- The File Manager treats the specified folder as the root level.
- Only the contents of that folder are displayed.
- Navigation, breadcrumbs, and file operations are scoped within this nested root.

Leave `RootFolderId` empty to use the OneDrive drive root as the File Manager root.

## How to Run This Application

Clone this repository to your local machine:

```bash
git clone https://github.com/SyncfusionExamples/onedrive-aspcore-file-provider
cd onedrive-aspcore-file-provider/onedrive-aspcore-file-provider
```

### Restore NuGet Packages

```bash
dotnet restore
```

### Build the Application

```bash
dotnet build
```

### Run the Application

```bash
dotnet run
```

The project will be hosted on `http://localhost`. To ensure the OneDrive service is running, navigate to:

```
http://localhost:<port-number>/api/test
```

## File Manager AjaxSettings

To access the basic actions such as Read, Delete, Copy, Move, Rename, Search, and Get Details of File Manager using the OneDrive service, map the following code snippet in the `AjaxSettings` property of the File Manager. Here, the `hostUrl` will be your locally hosted port number.

```javascript
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations'
}
```

### File Download AjaxSettings

To perform the download operation, initialize the `downloadUrl` property in `ajaxSettings` of the File Manager component.

```javascript
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    downloadUrl: hostUrl + 'api/OneDriveProvider/OneDriveDownload'
}
```

### File Upload AjaxSettings

To perform the upload operation, initialize the `uploadUrl` property in `ajaxSettings` of the File Manager component.

```javascript
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    uploadUrl: hostUrl + 'api/OneDriveProvider/OneDriveUpload'
}
```

### File Image Preview AjaxSettings

To enable image preview support in the File Manager component, initialize the `getImageUrl` property in `ajaxSettings` of the File Manager component.

```javascript
var hostUrl = http://localhost:62870/;
ajaxSettings: {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    getImageUrl: hostUrl + 'api/OneDriveProvider/OneDriveGetImage'
}
```

The FileManager will be rendered as the following.

![File Manager](https://ej2.syncfusion.com/products/images/file-manager/readme.gif)

## Support

Product support is available through the following mediums:

* Get expert help when you need it by [contacting support](https://www.syncfusion.com/support/directtrac/incidents?utm_source=npm&utm_campaign=filemanager), or ask, share, and connect with peers in the [community forum](https://www.syncfusion.com/forums/essential-js2?utm_source=npm&utm_campaign=filemanager).

* New [GitHub issue](https://github.com/syncfusion/ej2-javascript-ui-controls/issues/new).

* Ask your query in [Stack Overflow](https://stackoverflow.com/?utm_source=npm&utm_campaign=filemanager) with tag `syncfusion` and `ej2`.

## License

Check the license details [here](https://github.com/syncfusion/ej2-javascript-ui-controls/blob/master/license).

## Changelog

Check the changelog [here](https://github.com/syncfusion/ej2-javascript-ui-controls/blob/master/controls/filemanager/CHANGELOG.md)

© Copyright 2026 Syncfusion, Inc. All Rights Reserved. The Syncfusion Essential Studio license and copyright applies to this distribution.
