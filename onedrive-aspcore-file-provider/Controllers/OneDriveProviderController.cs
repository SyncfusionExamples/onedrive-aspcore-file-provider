using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OneDriveASPCoreFileProvider.Models;
using Syncfusion.Web.FileManager.Base;
using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;

namespace OneDriveASPCoreFileProvider.Controllers
{
    [Route("api/[controller]")]
    [EnableCors("AllowAllOrigins")]
    public class OneDriveProviderController : Controller
    {
        public OneDriveProvider operation;

        public OneDriveProviderController(IConfiguration configuration)
        {
            this.operation = new OneDriveProvider();
            var oneDriveSettings = configuration.GetSection("OneDriveSettings");
            // Validate configuration at startup so misconfiguration fails fast with a clear error
            this.operation.RegisterOneDrive(oneDriveSettings["TenantId"], oneDriveSettings["ClientId"], oneDriveSettings["ClientSecret"], oneDriveSettings["DriveId"]);
            this.operation.driveId = oneDriveSettings["DriveId"];
            this.operation.rootFolderId = oneDriveSettings["RootFolderId"];
        }

        [Route("OneDriveFileOperations")]
        public object FileOperations([FromBody] FileManagerDirectoryContent args)
        {
            this.operation.AccessDetails = GetRules();
            if (args.Action == "delete" || args.Action == "rename")
            {
                // Block modification only when the *root* folder itself is the target.
                bool isRootPath = string.IsNullOrEmpty(args.Path) || args.Path == "/";
                bool isRootItem = false;
                if (args.Data != null && args.Data.Length > 0)
                {
                    var item = args.Data[0];
                    // If RootFolderId is configured, block by matching ID.
                    // If no RootFolderId is set, block only when the item has no parent (drive root).
                    if (!string.IsNullOrEmpty(this.operation.rootFolderId))
                    {
                        isRootItem = item.Id == this.operation.rootFolderId;
                    }
                    else
                    {
                        isRootItem = string.IsNullOrEmpty(item.ParentId);
                    }
                }

                if ((args.TargetPath == null) && isRootPath && isRootItem)
                {
                    FileManagerResponse response = new FileManagerResponse();
                    response.Error = new ErrorDetails { Code = "401", Message = "Restricted to modify the root folder." };
                    return this.operation.ToCamelCase(response);
                }
            }
            switch (args.Action)
            {
                case "read":
                    return this.operation.ToCamelCase(this.operation.GetFiles(args.Path, args.ShowHiddenItems, args.Data));
                case "delete":
                    return this.operation.ToCamelCase(this.operation.Delete(args.Path, args.Names, args.Data));
                case "copy":
                    return this.operation.ToCamelCase(this.operation.Copy(args.Action, args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData, args.Data));
                case "move":
                    return this.operation.ToCamelCase(this.operation.Move(args.Action, args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData, args.Data));
                case "details":
                    return this.operation.ToCamelCase(this.operation.Details(args.Path, args.Names, args.Data));
                case "create":
                    return this.operation.ToCamelCase(this.operation.Create(args.Path, args.Name, args.Data));
                case "search":
                    return this.operation.ToCamelCase(this.operation.Search(args.Path, args.SearchString, args.ShowHiddenItems, args.CaseSensitive, args.FilterType, args.Data));
                case "rename":
                    return this.operation.ToCamelCase(this.operation.Rename(args.Path, args.Name, args.NewName, false, args.ShowFileExtension, args.Data));
            }
            return null;
        }

        [Route("OneDriveUpload")]
        [DisableRequestSizeLimit]
        public IActionResult OneDriveUpload(string path, IList<IFormFile> uploadFiles, string action, string data, long size = 0)
        {
            FileManagerResponse uploadResponse;
            FileManagerDirectoryContent[] dataObject = new FileManagerDirectoryContent[1];
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = null,
                PropertyNameCaseInsensitive = true
            };
            dataObject[0] = JsonSerializer.Deserialize<FileManagerDirectoryContent>(data, options);
            // The route value for 'action' (the action name 'OneDriveUpload') can shadow the
            // client-supplied upload action (save/replace/keepboth/remove) from the request.
            // Prefer the explicit upload action sent by the File Manager.
            string uploadAction = Request.HasFormContentType && Request.Form != null && Request.Form.ContainsKey("action")
                ? Request.Form["action"].ToString()
                : Request.Query["action"].ToString();
            if (string.IsNullOrEmpty(uploadAction) || uploadAction == "OneDriveUpload")
            {
                uploadAction = action;
            }
            uploadResponse = this.operation.Upload(path, uploadFiles, uploadAction, dataObject, size);
            if (uploadResponse.Error != null)
            {
                Response.Clear();
                Response.ContentType = "application/json; charset=utf-8";
                Response.StatusCode = Convert.ToInt32(uploadResponse.Error.Code);
                Response.HttpContext.Features.Get<IHttpResponseFeature>().ReasonPhrase = uploadResponse.Error.Message;
            }
            return Content("");
        }

        [Route("OneDriveDownload")]
        public IActionResult Download(string downloadInput)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            };
            FileManagerDirectoryContent args = JsonSerializer.Deserialize<FileManagerDirectoryContent>(downloadInput, options);
            return this.operation.Download(args.Path, args.Names, args.Data);
        }

        [Route("OneDriveGetImage")]
        public IActionResult GetImage(FileManagerDirectoryContent args)
        {
            return this.operation.GetImage(args.Path, args.Id, false, args.Data);
        }

        public AccessDetails GetRules()
        {
            AccessDetails accessDetails = new AccessDetails();
            List<AccessRule> Rules = new List<AccessRule>
            {
                new AccessRule { Path = "Files", Role = "Document Manager", Copy = Permission.Allow, Read = Permission.Allow, Write = Permission.Allow, WriteContents = Permission.Allow, Download = Permission.Allow, Upload = Permission.Allow, IsFile = false },
                new AccessRule { Path = "Microsoft Copilot Chat Files", Role = "Document Manager", Copy = Permission.Deny, Read = Permission.Allow, Write = Permission.Deny, WriteContents = Permission.Deny, Download = Permission.Allow, Upload = Permission.Deny, IsFile = false },
                new AccessRule { Path = "Apps", Role = "Document Manager", Copy = Permission.Deny, Read = Permission.Deny, Write = Permission.Deny, WriteContents = Permission.Deny, Download = Permission.Deny, Upload = Permission.Deny, IsFile = false }
            };
            accessDetails.AccessRules = Rules;
            accessDetails.Role = "Document Manager";
            return accessDetails;
        }
    }
}
