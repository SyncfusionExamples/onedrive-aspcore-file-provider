# onedrive-aspcore-file-provider

This repository contains the **OneDrive file system provider** in ASP.NET Core for the Syncfusion [File Manager component](https://www.syncfusion.com/ej2-asp-core-file-manager). It enables the File Manager to browse, manage, and perform file operations directly against a Microsoft OneDrive drive through the [Microsoft Graph API](https://developer.microsoft.com/en-us/graph).

## Key Features

The following actions can be performed using the OneDrive file system provider:

| **Actions** | **Description** |
| --- | --- |
| **Read** | Reads the files and folders from OneDrive, including navigation breadcrumbs and item metadata. |
| **Details** | Gets a file or folder's details such as Type, Size, Location, and Modified date. |
| **Download** | Downloads the selected file or folder as a ZIP archive from OneDrive. |
| **Upload** | Uploads one or more files to OneDrive. Supports drag-and-drop and request-size limits. |
| **Create** | Creates a new folder at the selected location. |
| **Delete** | Deletes the selected folders or files. |
| **Copy** | Copies the selected files or folders to a target location, with conflict handling. |
| **Move** | Moves the selected files or folders to the desired location, with conflict handling. |
| **Rename** | Renames a folder or file. |
| **Search** | Searches for files and folders using a text query within the current context. |

## Prerequisites

This provider must be registered as an Azure Active Directory application and granted the following Microsoft Graph **application permissions** with admin consent:

- **Files.Read.All**
- **Files.ReadWrite.All**

For details on registering an application, see [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).

## Configuration

Open `appsettings.json` and set your OneDrive account details under the `OneDriveSettings` section:

```json
{
  "OneDriveSettings": {
    "TenantId": "<--Tenant Id-->",
    "ClientId": "<--Client Id-->",
    "ClientSecret": "<--Client Secret-->",
    "DriveId": "<--Drive Id (optional)-->",
    "RootFolderId": ""
  }
}
```

| Setting | Description |
|---|---|
| `TenantId` | Azure AD tenant ID of the registered application. |
| `ClientId` | Application (client) ID of the registered application. |
| `ClientSecret` | Client secret value created for the registered application. |
| `DriveId` | Optional. The OneDrive drive ID to use. If omitted, the default drive is selected automatically. |
| `RootFolderId` | Optional. The OneDrive item ID of a specific folder to scope the File Manager root. When set, the File Manager will display only this folder and its contents. |

## Access Rules

The provider contains a simple, built-in access rule example. You can customize these rules in `Controllers/OneDriveProviderController.cs` inside the `GetRules()` method.

Three default rules are configured:

1. **Files** - Full access (read, write, copy, download, upload).
2. **Microsoft Copilot Chat Files** - Read-only access (read and download only).
3. **Apps** - Hidden and inaccessible to users.

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

By default, the application listens on `http://0.0.0.0:62870` and can also be accessed via `http://localhost:62870`.

To verify the service is running, navigate to:

```
http://localhost:62870/api/test
```

## File Manager AjaxSettings

Configure the File Manager component's `ajaxSettings` to point to the provider endpoints. Replace `<port-number>` with the actual port used by the provider.

If the controller name in your solution is `OneDriveProviderController`, the route resolves to:

```
api/OneDriveProvider/<action-name>
```

### Basic File Operations

```javascript
var hostUrl = 'http://localhost:62870/';
var ajaxSettings = {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations'
};
```

### With Download Support

```javascript
var hostUrl = 'http://localhost:62870/';
var ajaxSettings = {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    downloadUrl: hostUrl + 'api/OneDriveProvider/OneDriveDownload'
};
```

### With Upload Support

```javascript
var hostUrl = 'http://localhost:62870/';
var ajaxSettings = {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    uploadUrl: hostUrl + 'api/OneDriveProvider/OneDriveUpload'
};
```

### With Image Preview Support

```javascript
var hostUrl = 'http://localhost:62870/';
var ajaxSettings = {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    getImageUrl: hostUrl + 'api/OneDriveProvider/OneDriveGetImage'
};
```

### Complete AjaxSettings Example

```javascript
var hostUrl = 'http://localhost:62870/';
var ajaxSettings = {
    url: hostUrl + 'api/OneDriveProvider/OneDriveFileOperations',
    downloadUrl: hostUrl + 'api/OneDriveProvider/OneDriveDownload',
    uploadUrl: hostUrl + 'api/OneDriveProvider/OneDriveUpload',
    getImageUrl: hostUrl + 'api/OneDriveProvider/OneDriveGetImage'
};
```

The provider exposes the following endpoints:

| Endpoint | Action |
|---|---|
| `api/OneDriveProvider/OneDriveFileOperations` | Read, create, delete, copy, move, rename, details, search. |
| `api/OneDriveProvider/OneDriveUpload` | Upload files. |
| `api/OneDriveProvider/OneDriveDownload` | Download selected files or folders as a ZIP. |
| `api/OneDriveProvider/OneDriveGetImage` | Retrieve image content for preview. |

## Running the File Manager

After the application is running and the File Manager component is configured with the AjaxSettings above, bind the component in your EJ2 or Blazor application:

```html
<ejs-filemanager id="file-manager" ajaxSettings="ajaxSettings"></ejs-filemanager>
```

The File Manager will render and allow end users to interact with their OneDrive files.

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
