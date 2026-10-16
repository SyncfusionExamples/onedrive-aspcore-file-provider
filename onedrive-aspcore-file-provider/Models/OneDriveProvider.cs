using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.Copy;
using Microsoft.Graph.Models;
using Syncfusion.Web.FileManager.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Threading.Tasks;

namespace OneDriveASPCoreFileProvider.Models
{
    public class OneDriveProvider
    {
        public TokenCredential _tokenCredential { get; set; } = default;
        public string[] _scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
        public GraphServiceClient graphServiceClient { get; set; } = default;
        public string driveId { get; set; } = string.Empty;
        public string rootFolderId { get; set; } = string.Empty;
        public AccessDetails AccessDetails { get; set; } = new AccessDetails();
        private string accessMessage { get; set; } = string.Empty;

        public void RegisterOneDrive(string tenantId, string clientId, string clientSecret, string driveId)
        {
            if (string.IsNullOrWhiteSpace(tenantId))
            {
                throw new ArgumentNullException(nameof(tenantId), "OneDriveSettings:TenantId is missing in appsettings.json. The provider uses app-only authentication and requires a valid Azure AD tenant ID.");
            }
            if (string.IsNullOrWhiteSpace(clientId))
            {
                throw new ArgumentNullException(nameof(clientId), "OneDriveSettings:ClientId is missing in appsettings.json. The provider uses app-only authentication and requires a valid Azure AD application (client) ID.");
            }
            if (string.IsNullOrWhiteSpace(clientSecret))
            {
                throw new ArgumentNullException(nameof(clientSecret), "OneDriveSettings:ClientSecret is missing in appsettings.json. The provider uses app-only authentication and requires a valid client secret.");
            }
            if (string.IsNullOrWhiteSpace(driveId))
            {
                throw new ArgumentNullException(nameof(driveId), "OneDriveSettings:DriveId is missing in appsettings.json. The provider uses app-only (client credentials) authentication, so an explicit drive ID is required. The /me endpoint cannot be used to resolve a drive for an application principal.");
            }
            _tokenCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            graphServiceClient = new GraphServiceClient(_tokenCredential, _scopes);
        }

        public async Task<DriveItem> GetRootItemAsync()
        {
            if (!string.IsNullOrEmpty(rootFolderId))
            {
                return await graphServiceClient.Drives[driveId].Items[rootFolderId].GetAsync();
            }
            return await graphServiceClient.Drives[driveId].Root.GetAsync();
        }

        public FileManagerResponse GetFiles(string path, bool showHiddenItems, params FileManagerDirectoryContent[] data)
        {
            return GetFilesAsync(path, showHiddenItems, data).GetAwaiter().GetResult();
        }

        public async Task<FileManagerResponse> GetFilesAsync(string path, bool showHiddenItems, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse readResponse = new FileManagerResponse();
            try
            {
                DriveItem root = null;
                
                try
                {
                    if (path == "/")
                    {
                        root = await GetRootItemAsync();
                    }
                    else if (data != null && data.Length > 0 && !string.IsNullOrEmpty(data[0].Id))
                    {
                        // Primary: Use ID from File Manager request (fastest) - enforce root boundary
                        root = await EnsureItemWithinRootAsync(data[0].Id);
                    }
                    else if (!string.IsNullOrEmpty(path) && path != "/" && IsGraphItemId(path))
                    {
                        // Secondary: Detect if path itself is a Graph ID (e.g., from persistence/breadcrumb) - enforce root boundary
                        root = await EnsureItemWithinRootAsync(path.Trim('/'));
                    }
                    else
                    {
                        // Fallback: Resolve name-based path to folder (slowest, but necessary) - enforce root boundary
                        root = await ResolveFolderPathAsync(path, await GetRootItemAsync());
                    }
                    // Enforce root boundary for the resolved item (covers all resolution paths).
                    if (root != null && !await IsItemUnderRoot(root))
                    {
                        throw new UnauthorizedAccessException("The requested folder is outside the configured root folder boundary.");
                    }
                }
                catch
                {
                    throw;
                }

                // Safety check: Ensure we got a valid root
                if (root == null)
                {
                    throw new InvalidOperationException("Failed to fetch root folder - got null response from Graph API");
                }

                FileManagerDirectoryContent cwd = await ToFileManagerDirectoryContent(root);
                
                // Safety check: Ensure CWD is valid
                if (cwd == null)
                {
                    throw new InvalidOperationException("Failed to convert root folder to FileManagerDirectoryContent - got null response");
                }
                
                readResponse.CWD = cwd;

                if (cwd.Permission != null && !cwd.Permission.Read)
                {
                    readResponse.Files = null;
                    throw new UnauthorizedAccessException("'" + cwd.Name + "' is not accessible. You need permission to perform the read action.");
                }

                var children = await GetChildAsync(root.Id);
                readResponse.Files = children ?? new List<FileManagerDirectoryContent>();  // Ensure never null
                return readResponse;
            }
            catch (Exception e)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = e.Message;
                er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                readResponse.Error = er;
                return readResponse;
            }
        }

        private async Task<DriveItem> ResolveFolderPathAsync(string path, DriveItem root)
        {
            var pathSegments = (path ?? string.Empty)
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            if (pathSegments.Length == 0)
            {
                return root;
            }

            var firstSegment = pathSegments[0];
            if (string.Equals(firstSegment, root.Name, StringComparison.OrdinalIgnoreCase))
            {
                pathSegments = pathSegments.Skip(1).ToArray();
            }

            var currentFolder = root;
            foreach (var segment in pathSegments)
            {
                if (currentFolder.Folder == null)
                {
                    throw new DirectoryNotFoundException($"'{currentFolder.Name}' is not a folder in path '{path}'.");
                }

                DriveItem child = null;
                var childrenRequest = graphServiceClient.Drives[driveId].Items[currentFolder.Id].Children;
                var children = await childrenRequest.GetAsync();
                while (children != null)
                {
                    child = children.Value?.FirstOrDefault(item =>
                        string.Equals(item.Name, segment, StringComparison.OrdinalIgnoreCase));
                    if (child != null || string.IsNullOrEmpty(children.OdataNextLink))
                    {
                        break;
                    }

                    children = await childrenRequest.WithUrl(children.OdataNextLink).GetAsync();
                }

                if (child == null)
                {
                    throw new DirectoryNotFoundException($"Folder '{segment}' was not found in path '{path}'.");
                }

                if (child.Folder == null)
                {
                    throw new DirectoryNotFoundException($"'{segment}' is not a folder in path '{path}'.");
                }

                currentFolder = child;
            }

            return currentFolder;
        }

        public FileManagerResponse Create(string path, string name, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return CreateAsync(path, name, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> CreateAsync(string path, string name, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse createResponse = new FileManagerResponse();
            try
            {
                string parentId = data?.FirstOrDefault()?.Id ?? (await GetRootItemAsync())?.Id;
                // Enforce root boundary on the create destination
                await EnsureItemWithinRootAsync(parentId);
                AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, false);

                var targetChildren = await GetAllChildrenAsync(parentId);
                if (targetChildren.Any(targetItem => targetItem.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    createResponse.Error = new ErrorDetails
                    {
                        Code = "400",
                        Message = "A file or folder with the name " + name + " already exists."
                    };
                    return createResponse;
                }

                if (PathPermission != null && (!PathPermission.Read || !PathPermission.WriteContents))
                {
                    accessMessage = PathPermission.Message;
                    throw new UnauthorizedAccessException("'" + name + "' is not accessible. You need permission to perform the writeContents action.");
                }

                var folder = new DriveItem
                {
                    Name = name,
                    Folder = new Folder()
                };
                await graphServiceClient.Drives[driveId].Items[parentId].Children.PostAsync(folder);
                var createdItem = await graphServiceClient.Drives[driveId].Items[parentId].Children[name].GetAsync();
                createResponse.Files = new List<FileManagerDirectoryContent> { await ToFileManagerDirectoryContent(createdItem) };
                return createResponse;
            }
            catch (Exception ex)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = ex.Message;
                er.Code = ex.Message.Contains("is not accessible. You need permission") ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                createResponse.Error = er;
                return createResponse;
            }
        }

        public FileManagerResponse Delete(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return AsyncDelete(path, names, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> AsyncDelete(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse removeResponse = new FileManagerResponse();
            try
            {
                List<FileManagerDirectoryContent> deletedItems = new List<FileManagerDirectoryContent>();
                foreach (FileManagerDirectoryContent item in data)
                {
                    // Enforce root boundary before deleting a request-supplied item ID
                    await EnsureItemWithinRootAsync(item.Id);
                    AccessPermission PathPermission = GetPathPermission(item.FilterPath + item.Name, item.IsFile);
                    if (PathPermission != null && (!PathPermission.Read || !PathPermission.Write))
                    {
                        accessMessage = PathPermission.Message;
                        throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible.  You need permission to perform the write action.");
                    }
                    await graphServiceClient.Drives[driveId].Items[item.Id].DeleteAsync();
                    deletedItems.Add(item);
                }
                removeResponse.Files = deletedItems;
                return removeResponse;
            }
            catch (Exception ex)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = ex.Message;
                er.Code = ex.Message.Contains(" is not accessible.  You need permission") ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                removeResponse.Error = er;
                return removeResponse;
            }
        }

        public FileManagerResponse Rename(string path, string name, string newName, bool replace = false, bool showFileExtension = true, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return RenameAsync(path, name, newName, replace, showFileExtension, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> RenameAsync(string path, string name, string newName, bool replace, bool showFileExtension, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse renameResponse = new FileManagerResponse();
            try
            {
                // Enforce root boundary before renaming a request-supplied item ID
                await EnsureItemWithinRootAsync(data[0].Id);
                AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, data[0].IsFile);
                if (PathPermission != null && (!PathPermission.Read || !PathPermission.Write))
                {
                    accessMessage = PathPermission.Message;
                    throw new UnauthorizedAccessException();
                }

                string parentId = data[0].ParentId;
                if (string.IsNullOrEmpty(parentId))
                {
                    var item = await graphServiceClient.Drives[driveId].Items[data[0].Id].GetAsync();
                    parentId = item?.ParentReference?.Id;
                    if (string.IsNullOrEmpty(parentId) && !string.IsNullOrEmpty(rootFolderId) && data[0].Id == rootFolderId)
                    {
                        parentId = rootFolderId;
                    }
                }
                var targetChildren = await GetAllChildrenAsync(parentId);
                if (targetChildren.Any(targetItem => targetItem.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
                {
                    renameResponse.Error = new ErrorDetails
                    {
                        Code = "400",
                        Message = "Cannot rename " + name + " to " + newName + ": destination already exists."
                    };
                    return renameResponse;
                }

                var renameItem = new DriveItem { Name = newName };
                await graphServiceClient.Drives[driveId].Items[data[0].Id].PatchAsync(renameItem);
                var renamedItem = await graphServiceClient.Drives[driveId].Items[data[0].Id].GetAsync();
                renameResponse.Files = new List<FileManagerDirectoryContent> { await ToFileManagerDirectoryContent(renamedItem) };
                return renameResponse;
            }
            catch (Exception ex)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = (ex.GetType().Name == "UnauthorizedAccessException") ? "'" + name + "' is not accessible. You need permission to perform the write action." : ex.Message;
                er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                renameResponse.Error = er;
                return renameResponse;
            }
        }

        public FileManagerResponse Copy(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return CopyOrMoveDriveItemAsync(action, path, targetPath, names, replacedItemNames, TargetData, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public FileManagerResponse Move(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return CopyOrMoveDriveItemAsync(action, path, targetPath, names, replacedItemNames, TargetData, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> CopyOrMoveDriveItemAsync(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse moveResponse = new FileManagerResponse();
            List<FileManagerDirectoryContent> transferredList = new List<FileManagerDirectoryContent>();
            try
            {
                if (data.Length == 0) {
                    moveResponse.Files = new List<FileManagerDirectoryContent>();
                    return moveResponse;
                }

                // Enforce root boundary on every source item (do not trust client-supplied FilterPath/ParentId)
                foreach (FileManagerDirectoryContent sourceItem in data)
                {
                    await EnsureItemWithinRootAsync(sourceItem.Id);
                }

                AccessPermission sourcePermission = GetPathPermission(data[0].FilterPath + data[0].Name, data[0].IsFile);
                if (action == "copy")
                {
                    if (sourcePermission != null && (!sourcePermission.Read || !sourcePermission.Copy))
                    {
                        accessMessage = sourcePermission.Message;
                        throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the copy action.");
                    }
                }
                else
                {
                    if (sourcePermission != null && (!sourcePermission.Read || !sourcePermission.Write))
                    {
                        accessMessage = sourcePermission.Message;
                        throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the Write action.");
                    }
                }

                if (string.IsNullOrEmpty(TargetData?.Id))
                {
                    moveResponse.Error = new ErrorDetails { Code = "404", Message = "Target location is not available." };
                    return moveResponse;
                }

                // Enforce root boundary on the copy/move destination as well
                await EnsureItemWithinRootAsync(TargetData.Id);

                var targetItems = await GetAllChildrenAsync(TargetData.Id);

                List<string> existFiles = new List<string>();
                List<FileManagerDirectoryContent> otherFiles = new List<FileManagerDirectoryContent>();

                foreach (FileManagerDirectoryContent item in data)
                {
                    if (targetItems.Any(t => t.Name.Equals(item.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        existFiles.Add(item.Name);
                    }
                    else
                    {
                        otherFiles.Add(item);
                    }
                }

                if ((replacedItemNames == null || replacedItemNames.Length == 0) && existFiles.Count > 0)
                {
                    ErrorDetails er = new ErrorDetails();
                    er.FileExists = existFiles;
                    er.Code = "400";
                    er.Message = "File Already Exists";
                    moveResponse.Files = otherFiles;
                    moveResponse.Error = er;
                    return moveResponse;
                }

                var targetNamesUsed = new List<string>();
                var itemsToRefresh = new List<DriveItem>();

                foreach (FileManagerDirectoryContent item in data)
                {
                    var existingTarget = targetItems.FirstOrDefault(t => t.Name.Equals(item.Name, StringComparison.OrdinalIgnoreCase));
                    bool conflict = existingTarget != null;

                    // Syncfusion conflict-dialog contract: 'RenameFiles' (replacedItemNames) contains
                    // the source names the user chose to RENAME (keep both). A conflicting name that
                    // is NOT in that list means the user chose REPLACE: the existing target item is
                    // deleted and the source takes its place.
                    bool renameThis = conflict && replacedItemNames != null && replacedItemNames.Contains(item.Name);

                    string targetName = item.Name;
                    if (conflict)
                    {
                        if (renameThis)
                        {
                            // User chose "keep both" (rename): copy/move under a unique name.
                            targetName = await GenerateUniqueTargetNameAsync(TargetData.Id, item.Name);
                        }
                        else
                        {
                            // User chose "replace": remove the existing target item first so the
                            // source item can take its exact name.
                            await graphServiceClient.Drives[driveId].Items[existingTarget.Id].DeleteAsync();
                        }
                    }

                    if (action == "copy")
                    {
                        var requestBody = new CopyPostRequestBody
                        {
                            ParentReference = new ItemReference { DriveId = driveId, Id = TargetData.Id },
                            Name = targetName
                        };
                        var operation = await graphServiceClient.Drives[driveId].Items[item.Id].Copy.PostAsync(requestBody);

                        var copiedItem = await WaitForCopiedItemAsync(TargetData.Id, targetName);
                        if (copiedItem != null)
                        {
                            itemsToRefresh.Add(copiedItem);
                        }
                    }
                    else
                    {
                        var moveRequestBody = new DriveItem
                        {
                            ParentReference = new ItemReference { Id = TargetData.Id },
                            Name = targetName
                        };
                        await graphServiceClient.Drives[driveId].Items[item.Id].PatchAsync(moveRequestBody);

                        var movedItem = await graphServiceClient.Drives[driveId].Items[item.Id].GetAsync();
                        if (movedItem != null)
                        {
                            itemsToRefresh.Add(movedItem);
                        }
                    }

                    targetNamesUsed.Add(targetName);
                }

                if (itemsToRefresh.Any())
                {
                    foreach (var driveItem in itemsToRefresh)
                    {
                        transferredList.Add(await ToFileManagerDirectoryContent(driveItem));
                    }
                }
                moveResponse.Files = transferredList;
                return moveResponse;
            }
            catch (Exception ex)
            {
                ErrorDetails error = new ErrorDetails();
                error.Message = ex.Message;
                error.Code = ex.Message.Contains("is not accessible. You need permission") ? "401" : "404";
                if (error.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { error.Message = accessMessage; }
                moveResponse.Error = error;
                return moveResponse;
            }
        }

        private async Task<string> GenerateUniqueTargetNameAsync(string parentId, string name)
        {
            string extension = Path.GetExtension(name);
            string baseName = Path.GetFileNameWithoutExtension(name);
            if (string.IsNullOrEmpty(baseName)) baseName = name;
            int count = 1;
            string newName;
            do
            {
                newName = $"{baseName}({count}){extension}";
                count++;
            } while (await ItemExistsAtTargetAsync(parentId, newName));
            return newName;
        }

        private async Task<bool> ItemExistsAtTargetAsync(string parentId, string name)
        {
            try
            {
                var child = await graphServiceClient.Drives[driveId].Items[parentId].ItemWithPath(name).GetAsync();
                return child != null;
            }
            catch
            {
                return false;
            }
        }

        private async Task<DriveItem> WaitForCopiedItemAsync(string parentId, string name, int maxAttempts = 30)
        {
            int attempts = 0;
            while (attempts < maxAttempts)
            {
                try
                {
                    var child = await graphServiceClient.Drives[driveId].Items[parentId].ItemWithPath(name).GetAsync();
                    if (child != null)
                    {
                        return child;
                    }
                }
                catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
                {
                    // 404 / item not yet available, keep waiting
                    if (ex.ResponseStatusCode != 404)
                    {
                    }
                }
                catch
                {
                }
                attempts++;
                await Task.Delay(1000);
            }
            return null;
        }

        private async Task<List<FileManagerDirectoryContent>> WaitForTransferredItemsAsync(string parentId, List<string> names, int maxAttempts = 10)
        {
            List<FileManagerDirectoryContent> result = new List<FileManagerDirectoryContent>();
            foreach (string name in names)
            {
                DriveItem transferredDrive = await WaitForCopiedItemAsync(parentId, name, maxAttempts);
                if (transferredDrive != null)
                {
                    result.Add(await ToFileManagerDirectoryContent(transferredDrive));
                }
            }
            return result;
        }

        public FileManagerResponse Details(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return DetailsAsync(path, names, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> DetailsAsync(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse detailsResponse = new FileManagerResponse();
            FileDetails detailFiles = new FileDetails();
            try
            {
                // Enforce root boundary on every item being detailed
                foreach (var item in data)
                {
                    await EnsureItemWithinRootAsync(item.Id);
                }
                if (data.Length == 0 || data.Length == 1)
                {
                    var root = data.FirstOrDefault();
                    detailFiles.Name = root.Name;
                    detailFiles.IsFile = root.IsFile;
                    detailFiles.Size = ByteConversion(root.Size);
                    detailFiles.Created = root.DateCreated;
                    detailFiles.Location = (root.ParentId == null || root.ParentId == "/") ? "root" : "root" + root.FilterPath.TrimEnd('/');
                    detailFiles.Modified = root.DateModified;
                    detailFiles.Permission = root.Permission;
                    detailFiles.MultipleFiles = false;
                }
                else
                {
                    detailFiles.MultipleFiles = true;
                    detailFiles.Name = string.Join(", ", data.Select(x => x.Name).ToArray());
                    var first = data.FirstOrDefault();
                    detailFiles.Location = (first.ParentId == null || first.ParentId == "/") ? "root" : "root" + first.FilterPath.TrimEnd('/');
                    long sizeValue = 0;
                    foreach (var item in data) { sizeValue += item.Size; }
                    detailFiles.Size = ByteConversion(sizeValue);
                }
                detailsResponse.Details = detailFiles;
                return detailsResponse;
            }
            catch (Exception e)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = e.Message;
                er.Code = e.Message.Contains("is not accessible. You need permission") ? "401" : "417";
                detailsResponse.Error = er;
                return detailsResponse;
            }
        }

        public FileManagerResponse Search(string path, string searchString, bool showHiddenItems, bool caseSensitive, string filterType, params FileManagerDirectoryContent[] data)
        {
            try
            {
                return SearchAsync(path, searchString, showHiddenItems, caseSensitive, filterType, data).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public async Task<FileManagerResponse> SearchAsync(string path, string searchString, bool showHiddenItems, bool caseSensitive, string filterType, params FileManagerDirectoryContent[] data)
        {
            FileManagerResponse searchResponse = new FileManagerResponse();
            try
            {
                AccessPermission permission = GetPathPermission(data[0].FilterPath + (data[0].IsFile ? data[0].Name : Path.GetFileNameWithoutExtension(data[0].Name)), data[0].IsFile);
                if (permission != null && !permission.Read)
                {
                    searchResponse.Files = null;
                    accessMessage = data[0].Permission.Message;
                    throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the read action.");
                }

                var query = (searchString ?? string.Empty).Trim().Trim('*');
                var escapedQuery = Regex.Escape(query)
                    .Replace(@"\*", ".*")
                    .Replace(@"\?", ".");
                var searchPattern = (filterType ?? string.Empty).Trim().ToLowerInvariant() switch
                {
                    "startswith" => "^" + escapedQuery + ".*$",
                    "endswith" => "^.*" + escapedQuery + "$",
                    "contains" or "" => "^.*" + escapedQuery + ".*$",
                    _ => "^.*" + escapedQuery + ".*$"
                };
                var searchRegex = new Regex(
                    searchPattern,
                    caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);

                // Determine the folder to start the search from
                string searchStartId;
                if (data != null && data.Length > 0 && !string.IsNullOrEmpty(data[0].Id))
                {
                    searchStartId = data[0].Id;
                    // Enforce root boundary on the search start folder
                    await EnsureItemWithinRootAsync(searchStartId);
                }
                else
                {
                    var root = await GetRootItemAsync();
                    searchStartId = root?.Id;
                }

                // Match the wildcard pattern locally because Graph search does not support '*'.
                var files = new List<FileManagerDirectoryContent>();
                var searchErrors = new List<string>();
                await SearchChildrenRecursiveAsync(searchStartId, searchRegex, files, searchErrors);

                searchResponse.Files = files;
                searchResponse.CWD = data?.FirstOrDefault();
                if (searchErrors.Count > 0)
                {
                    // Explicitly represent partial results rather than returning an incomplete list silently.
                    searchResponse.Error = new ErrorDetails
                    {
                        Code = "417",
                        Message = "Search completed partially: " + string.Join("; ", searchErrors)
                    };
                }
                return searchResponse;
            }
            catch (Exception e)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = e.Message;
                er.Code = e.Message.Contains("is not accessible. You need permission") ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                searchResponse.Error = er;
                return searchResponse;
            }
        }

        private async Task<bool> IsItemUnderRoot(DriveItem item)
        {
            // No configured root boundary: every item on the drive is permitted.
            if (string.IsNullOrEmpty(rootFolderId))
            {
                return true;
            }
            // The configured root folder itself is within the boundary.
            if (item?.Id == rootFolderId)
            {
                return true;
            }
            var current = item;
            while (current?.ParentReference != null && !string.IsNullOrEmpty(current.ParentReference.Id))
            {
                if (current.ParentReference.Id == rootFolderId)
                {
                    return true;
                }
                try
                {
                    current = await graphServiceClient.Drives[driveId].Items[current.ParentReference.Id].GetAsync();
                }
                catch
                {
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// Shared root-boundary validation. Fetches the item (verifying it belongs to the
        /// configured drive) and ensures it is the configured RootFolderId or a descendant
        /// of it. Throws UnauthorizedAccessException when the item lies outside the root
        /// boundary. Must be called before any operation reads or modifies a request-supplied
        /// item ID (do not trust client-supplied Path/ParentId/FilterPath as proof of containment).
        /// </summary>
        private async Task<DriveItem> EnsureItemWithinRootAsync(string itemId)
        {
            var driveItem = await GetDriveItemWithPathCheckAsync(itemId);
            if (driveItem == null)
            {
                throw new UnauthorizedAccessException("The requested item does not exist on the drive.");
            }
            if (!await IsItemUnderRoot(driveItem))
            {
                throw new UnauthorizedAccessException("The requested item is outside the configured root folder boundary.");
            }
            return driveItem;
        }

        private async Task SearchChildrenRecursiveAsync(string parentId, Regex searchRegex, List<FileManagerDirectoryContent> results, List<string> searchErrors)
        {
            try
            {
                // Follow Graph continuation pages so large folders are searched completely.
                var items = await GetAllChildrenAsync(parentId);

                foreach (var item in items)
                {
                    if (item == null || string.IsNullOrEmpty(item.Name)) continue;

                    // Skip the virtual root itself
                    if (!string.IsNullOrEmpty(rootFolderId) && item.Id == rootFolderId) continue;

                    if (searchRegex.IsMatch(item.Name))
                    {
                        try
                        {
                            results.Add(await ToFileManagerDirectoryContent(item));
                        }
                        catch (Exception ex)
                        {
                            // Record items that fail to convert instead of silently dropping them.
                            searchErrors.Add($"Failed to read item '{item.Name}': {ex.Message}");
                        }
                    }

                    // Recurse into folders
                    if (item.Folder != null)
                    {
                        await SearchChildrenRecursiveAsync(item.Id, searchRegex, results, searchErrors);
                    }
                }
            }
            catch (Exception ex)
            {
                // Record inaccessible folders so the response explicitly represents partial results.
                searchErrors.Add($"Failed to read folder '{parentId}': {ex.Message}");
            }
        }

        public FileManagerResponse Upload(string path, IList<IFormFile> uploadFiles, string action, FileManagerDirectoryContent[] data, long size = 0)
        {
            try
            {
                return AsyncUpload(path, uploadFiles, action, data, size).GetAwaiter().GetResult();
            }
            finally
            {
            }
        }

        public virtual async Task<FileManagerResponse> AsyncUpload(string path, IList<IFormFile> uploadFiles, string action, FileManagerDirectoryContent[] data, long size = 0)
        {
            FileManagerResponse uploadResponse = new FileManagerResponse();
            try
            {
                if (data == null || data.Length == 0 || string.IsNullOrEmpty(data[0]?.Id))
                {
                    throw new ArgumentException("The upload destination folder was not provided.");
                }

                // Enforce root boundary on the upload destination folder
                await EnsureItemWithinRootAsync(data[0].Id);

                AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, false);
                if (PathPermission != null && (!PathPermission.Read || !PathPermission.Upload))
                {
                    accessMessage = PathPermission.Message;
                    throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the upload action.");
                }
                string driveItemId = data[0].Id;
                List<string> existFiles = new List<string>();
                foreach (IFormFile file in uploadFiles)
                {
                    var relativePath = GetUploadPathParts(file.FileName);
                    if (string.IsNullOrWhiteSpace(relativePath.FileName))
                    {
                        uploadResponse.Error = new ErrorDetails
                        {
                            Code = "400",
                            Message = "The uploaded file name is invalid."
                        };
                        return uploadResponse;
                    }

                    string destinationFolderId = await EnsureUploadFolderPathAsync(driveItemId, relativePath.FolderNames);
                    string fileName = relativePath.FileName;
                    string chunkPath = GetChunkStagingPath(destinationFolderId, fileName);
                    if (action == "remove")
                    {
                        if (System.IO.File.Exists(chunkPath))
                        {
                            System.IO.File.Delete(chunkPath);
                        }
                        var removeResult = await RemoveItemFromDriveAsync(destinationFolderId, fileName);
                        if (!removeResult.Success)
                        {
                            uploadResponse.Error = new ErrorDetails
                            {
                                Code = "417",
                                Message = $"Failed to remove '{fileName}' from the drive: {removeResult.Message}"
                            };
                            return uploadResponse;
                        }
                        continue;
                    }

                    long expectedFileSize = size > 0 ? size : file.Length;
                    if (file.Length > expectedFileSize)
                    {
                        uploadResponse.Error = new ErrorDetails
                        {
                            Code = "400",
                            Message = $"The received upload data for '{fileName}' exceeds the declared file size."
                        };
                        return uploadResponse;
                    }

                    bool isChunkedRequest = size > file.Length;
                    if (isChunkedRequest)
                    {
                        if (action == "save" && !System.IO.File.Exists(chunkPath) &&
                            await ItemExistsInDrive(destinationFolderId, fileName))
                        {
                            existFiles.Add(fileName);
                            continue;
                        }

                        Directory.CreateDirectory(Path.GetDirectoryName(chunkPath));
                        await using (var stagingStream = new FileStream(
                            chunkPath, FileMode.Append, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                        {
                            await file.CopyToAsync(stagingStream);
                        }

                        var stagedLength = new FileInfo(chunkPath).Length;
                        if (stagedLength > expectedFileSize)
                        {
                            System.IO.File.Delete(chunkPath);
                            uploadResponse.Error = new ErrorDetails
                            {
                                Code = "400",
                                Message = $"The received upload data for '{fileName}' exceeds the declared file size."
                            };
                            return uploadResponse;
                        }

                        if (stagedLength < expectedFileSize)
                        {
                            continue;
                        }

                        if (action == "keepboth")
                        {
                            fileName = await GenerateUniqueFileName(destinationFolderId, fileName);
                        }

                        try
                        {
                            await using var stagedFile = new FileStream(chunkPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);
                            await UploadStreamToDriveAsync(destinationFolderId, fileName, stagedFile);
                        }
                        finally
                        {
                            if (System.IO.File.Exists(chunkPath))
                            {
                                System.IO.File.Delete(chunkPath);
                            }
                        }
                        continue;
                    }

                    if (action == "save")
                    {
                        if (await ItemExistsInDrive(destinationFolderId, fileName))
                        {
                            existFiles.Add(fileName);
                            continue;
                        }
                        await using var stream = file.OpenReadStream();
                        await UploadStreamToDriveAsync(destinationFolderId, fileName, stream);
                    }
                    else if (action == "replace")
                    {
                        await using var stream = file.OpenReadStream();
                        await UploadStreamToDriveAsync(destinationFolderId, fileName, stream);
                    }
                    else if (action == "keepboth")
                    {
                        string newFileName = await GenerateUniqueFileName(destinationFolderId, fileName);
                        await using var stream = file.OpenReadStream();
                        await UploadStreamToDriveAsync(destinationFolderId, newFileName, stream);
                    }
                }

                if (existFiles.Count != 0)
                {
                    uploadResponse.Error = new ErrorDetails
                    {
                        FileExists = existFiles,
                        Code = "400",
                        Message = "File Already Exists"
                    };
                }
                return uploadResponse;
            }
            catch (Exception ex)
            {
                string uploadName = data?.FirstOrDefault()?.Name ?? "the selected folder";
                ErrorDetails er = new ErrorDetails();
                er.Message = (ex.GetType().Name == "UnauthorizedAccessException") ? "'" + uploadName + "' is not accessible. You need permission to perform the upload action." : ex.Message;
                er.Code = ex.GetType().Name == "UnauthorizedAccessException" ? "401" : "417";
                if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
                uploadResponse.Error = er;
                return uploadResponse;
            }
        }

        private async Task UploadStreamToDriveAsync(string driveItemId, string fileName, Stream stream)
        {
            await graphServiceClient.Drives[driveId].Items[driveItemId].ItemWithPath(fileName).Content.PutAsync(stream);
        }

        private async Task<string> EnsureUploadFolderPathAsync(string parentId, IEnumerable<string> folderNames)
        {
            string currentFolderId = parentId;
            foreach (var folderName in folderNames)
            {
                var folder = await FindChildByNameAsync(currentFolderId, folderName);
                if (folder == null)
                {
                    folder = await graphServiceClient.Drives[driveId].Items[currentFolderId].Children.PostAsync(new DriveItem
                    {
                        Name = folderName,
                        Folder = new Folder()
                    });
                }

                if (folder?.Folder == null || string.IsNullOrEmpty(folder.Id))
                {
                    throw new InvalidOperationException($"'{folderName}' exists but is not a folder.");
                }

                currentFolderId = folder.Id;
            }

            return currentFolderId;
        }

        private async Task<DriveItem> FindChildByNameAsync(string parentId, string name)
        {
            var childrenRequest = graphServiceClient.Drives[driveId].Items[parentId].Children;
            var children = await childrenRequest.GetAsync();
            while (children != null)
            {
                var match = children.Value?.FirstOrDefault(item =>
                    string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                if (match != null || string.IsNullOrEmpty(children.OdataNextLink))
                {
                    return match;
                }

                children = await childrenRequest.WithUrl(children.OdataNextLink).GetAsync();
            }

            return null;
        }

        /// <summary>
        /// Enumerates every child item of a folder, following Graph continuation pages
        /// (@odata.nextLink) until the list is exhausted. Callers must not rely on a
        /// single Children.GetAsync() page because large folders are paginated.
        /// </summary>
        private async Task<List<DriveItem>> GetAllChildrenAsync(string parentId)
        {
            var allItems = new List<DriveItem>();
            var childrenRequest = graphServiceClient.Drives[driveId].Items[parentId].Children;
            var children = await childrenRequest.GetAsync();
            while (children != null)
            {
                if (children.Value != null)
                {
                    allItems.AddRange(children.Value);
                }
                if (string.IsNullOrEmpty(children.OdataNextLink))
                {
                    break;
                }
                children = await childrenRequest.WithUrl(children.OdataNextLink).GetAsync();
            }
            return allItems;
        }

        private string GetChunkStagingPath(string parentId, string fileName)
        {
            byte[] key = SHA256.HashData(Encoding.UTF8.GetBytes(parentId + "\0" + fileName));
            string stagingDirectory = Path.Combine(Path.GetTempPath(), "OneDriveASPCoreFileProvider", "ChunkUploads");
            return Path.Combine(stagingDirectory, Convert.ToHexString(key) + ".part");
        }

        private (string[] FolderNames, string FileName) GetUploadPathParts(string uploadFileName)
        {
            if (string.IsNullOrWhiteSpace(uploadFileName))
            {
                return (Array.Empty<string>(), string.Empty);
            }

            string normalizedPath = uploadFileName.Replace('\\', '/');
            if (normalizedPath.StartsWith("/", StringComparison.Ordinal) ||
                (normalizedPath.Length >= 2 && char.IsLetter(normalizedPath[0]) && normalizedPath[1] == ':'))
            {
                throw new UnauthorizedAccessException("Access denied for Directory-traversal");
            }

            string[] segments = normalizedPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length == 0 || segments.Any(segment => segment == "." || segment == ".."))
            {
                throw new UnauthorizedAccessException("Access denied for Directory-traversal");
            }

            string[] safeSegments = segments.Select(SanitizeFileName).ToArray();
            if (safeSegments.Any(string.IsNullOrWhiteSpace))
            {
                throw new ArgumentException("The uploaded file name is invalid.");
            }

            return (safeSegments.Take(safeSegments.Length - 1).ToArray(), safeSegments[safeSegments.Length - 1]);
        }

        private async Task<(bool Success, string Message)> RemoveItemFromDriveAsync(string driveItemId, string itemName)
        {
            try
            {
                if (!string.IsNullOrEmpty(rootFolderId) && driveItemId == rootFolderId)
                {
                    await graphServiceClient.Drives[driveId].Items[rootFolderId].ItemWithPath(itemName).DeleteAsync();
                }
                else
                {
                    await graphServiceClient.Drives[driveId].Items[driveItemId].ItemWithPath(itemName).DeleteAsync();
                }
                return (true, "Item deleted successfully");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return string.Empty;
            }
            string name = Path.GetFileName(fileName.Replace("\\", "/"));
            char[] invalidChars = Path.GetInvalidFileNameChars().Concat(new[] { ':', '*', '?', '"', '<', '>', '|' }).ToArray();
            foreach (char invalid in invalidChars)
            {
                name = name.Replace(invalid.ToString(), string.Empty);
            }
            return name.Trim();
        }

        private async Task<DriveItem> GetDriveItemWithPathCheckAsync(string itemId)
        {
            var driveItem = await graphServiceClient.Drives[driveId].Items[itemId].GetAsync();
            if (driveItem == null)
            {
                return null;
            }
            var parentReference = driveItem.ParentReference;
            if (parentReference == null || string.IsNullOrEmpty(parentReference.DriveId))
            {
                parentReference = await GetParentReferenceAsync(itemId);
            }
            if (parentReference != null && !string.IsNullOrEmpty(parentReference.DriveId) && parentReference.DriveId != driveId)
            {
                throw new UnauthorizedAccessException("The requested item belongs to a different drive.");
            }
            return driveItem;
        }

        private async Task<ItemReference> GetParentReferenceAsync(string itemId)
        {
            try
            {
                var item = await graphServiceClient.Drives[driveId].Items[itemId].GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = new[] { "parentReference" };
                });
                return item?.ParentReference;
            }
            catch
            {
                return null;
            }
        }

        public async Task<bool> ItemExistsInDrive(string driveItemId, string itemName)
        {
            return await FindChildByNameAsync(driveItemId, itemName) != null;
        }

        private async Task<string> GenerateUniqueFileName(string driveItemId, string name)
        {
            string extension = Path.GetExtension(name);
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(name);
            int count = 0;
            string newName = name;
            while (await ItemExistsInDrive(driveItemId, newName))
            {
                count++;
                newName = $"{nameWithoutExtension}({count}){extension}";
            }
            return newName;
        }

        public virtual FileStreamResult Download(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            return DownloadSelectedDriveItemsAsync(path, names, data).GetAwaiter().GetResult();
        }

        public async Task<FileStreamResult> DownloadSelectedDriveItemsAsync(string path, string[] names, params FileManagerDirectoryContent[] data)
        {
            try
            {
                if (data.Length == 1 && data[0].IsFile)
                {
                    AccessPermission pathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, data[0].IsFile);
                    if (pathPermission != null && (!pathPermission.Read || !pathPermission.Download))
                    {
                        throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. Access is denied.");
                    }
                    // Enforce root boundary before streaming content
                    await EnsureItemWithinRootAsync(data[0].Id);
                    Stream stream = await graphServiceClient.Drives[driveId].Items[data[0].Id].Content.GetAsync();
                    return new FileStreamResult(stream, "application/octet-stream") { FileDownloadName = data[0].Name };
                }

                // Build the ZIP in a temporary file instead of a MemoryStream so peak memory
                // does not grow with the archive size and stays bounded for concurrent downloads.
                // DeleteOnClose ensures the temp file is removed when the response stream is disposed.
                string zipPath = Path.Combine(Path.GetTempPath(), "OneDriveASPCoreFileProvider", "Downloads",
                    Guid.NewGuid().ToString("N") + ".zip");
                Directory.CreateDirectory(Path.GetDirectoryName(zipPath));
                var downloadStream = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 81920,
                    FileOptions.Asynchronous | FileOptions.DeleteOnClose);
                try
                {
                    using (var archive = new ZipArchive(downloadStream, ZipArchiveMode.Create, true))
                    {
                        foreach (var item in data)
                        {
                            AccessPermission pathPermission = GetPathPermission(item.FilterPath + (item.IsFile ? item.Name : Path.GetFileNameWithoutExtension(item.Name)), item.IsFile);
                            if (pathPermission != null && (!pathPermission.Read || !pathPermission.Download))
                            {
                                throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible. Access is denied.");
                            }
                            // Enforce root boundary before adding each item to the ZIP
                            var driveItem = await EnsureItemWithinRootAsync(item.Id);
                            if (driveItem.File != null)
                            {
                                await AddFileToZipAsync(archive, item.Id, driveItem.Name, driveItem.Name);
                            }
                            else if (driveItem.Folder != null)
                            {
                                await AddFolderToZipAsync(archive, driveItem, driveItem.Name);
                            }
                        }
                    }
                    // Hand over the temp file for streaming; the OS deletes it when the stream closes.
                    downloadStream.Seek(0, SeekOrigin.Begin);
                    return new FileStreamResult(downloadStream, "application/octet-stream")
                    {
                        FileDownloadName = "SelectedFiles.zip"
                    };
                }
                catch
                {
                    await downloadStream.DisposeAsync();
                    throw;
                }
            }
            catch
            {
                throw;
            }
        }

        private async Task AddFileToZipAsync(ZipArchive archive, string itemId, string pathInArchive, string fileName)
        {
            var entry = archive.CreateEntry(pathInArchive);
            using (var entryStream = entry.Open())
            {
                var stream = await graphServiceClient.Drives[driveId].Items[itemId].Content.GetAsync();
                await stream.CopyToAsync(entryStream);
            }
        }

        private async Task AddFolderToZipAsync(ZipArchive archive, DriveItem folderItem, string currentFolderPath)
        {
            archive.CreateEntry(currentFolderPath + "/");
            // Follow Graph continuation pages so large folders are zipped completely.
            var folderContents = await GetAllChildrenAsync(folderItem.Id);
            foreach (var item in folderContents)
            {
                string itemPath = Path.Combine(currentFolderPath, item.Name);
                if (item.File != null)
                {
                    await AddFileToZipAsync(archive, item.Id, itemPath, item.Name);
                }
                else if (item.Folder != null)
                {
                    await AddFolderToZipAsync(archive, item, itemPath);
                }
            }
        }

        public virtual FileStreamResult GetImage(string path, string id, bool allowCompress, params FileManagerDirectoryContent[] data)
        {
            return GetImageAsync(path, id, allowCompress, data).GetAwaiter().GetResult();
        }

        public async Task<FileStreamResult> GetImageAsync(string path, string id, bool allowCompress, params FileManagerDirectoryContent[] data)
        {
            // Enforce root boundary before streaming image content
            await EnsureItemWithinRootAsync(id);
            Stream stream = await graphServiceClient.Drives[driveId].Items[id].Content.GetAsync();
            return new FileStreamResult(stream, "application/octet-stream");
        }

        private async Task<FileManagerDirectoryContent> ToFileManagerDirectoryContent(DriveItem item)
        {
            var file = new FileManagerDirectoryContent();
            file.Name = item?.Name ?? string.Empty;
            file.Id = item?.Id ?? string.Empty;
            file.ParentId = item?.ParentReference?.Id;
            file.Size = item?.Size ?? 0;
            file.IsFile = item?.Folder == null;
            file.DateModified = item?.LastModifiedDateTime?.DateTime ?? DateTime.MinValue;
            file.DateCreated = item?.CreatedDateTime?.DateTime ?? DateTime.MinValue;
            file.HasChild = (item?.Folder != null) && await HasChild(item);
            file.Type = Path.GetExtension(item?.Name ?? string.Empty);
            file.FilterId = GetFilterIdFromParentReference(item);  // Optimized: zero API calls
            
            // Try optimized path parsing first, fall back to async if needed
            file.FilterPath = GetFilterPathFromParentReference(item);
            if (string.IsNullOrEmpty(file.FilterPath))
            {
                try
                {
                    file.FilterPath = await GetFilterPath(item);  // Fallback to original method
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Error] GetFilterPath failed: {ex.Message}. Using root path for item '{file.Name}'");
                    file.FilterPath = "/";  // Default to root if all else fails
                }
            }
            
            // CRITICAL FOR PERSISTENCE: Set Path to the item ID
            // This allows persistence to correctly reload the same item
            // Persistence stores the Path field, and when reloading, uses it as the identifier
            file.Path = file.Id;
            
            try
            {
                file.Permission = GetPathPermission(file.FilterPath + (file.IsFile ? file.Name : Path.GetFileNameWithoutExtension(file.Name)), file.IsFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] GetPathPermission failed: {ex.Message}");
                file.Permission = null;  // Allow permission to be null
            }
            
            return file;
        }

        private async Task<List<FileManagerDirectoryContent>> GetChildAsync(string itemId)
        {
            List<FileManagerDirectoryContent> child = new List<FileManagerDirectoryContent>();
            try
            {
                // Follow Graph continuation pages so large folders are listed completely.
                var children = await GetAllChildrenAsync(itemId);
                foreach (var item in children)
                {
                    try
                    {
                        var converted = await ToFileManagerDirectoryContent(item);
                        if (converted != null)
                        {
                            child.Add(converted);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Error] Failed to convert child item: {ex.Message}");
                        // Continue with next item instead of crashing
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] GetChildAsync failed: {ex.Message}");
                // Return empty list instead of null
            }
            return child;
        }

        private string ByteConversion(long fileSize)
        {
            try
            {
                string[] index = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
                if (fileSize == 0) return "0 " + index[0];
                long bytes = Math.Abs(fileSize);
                int loc = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
                double num = Math.Round(bytes / Math.Pow(1024, loc), 1);
                return (Math.Sign(fileSize) * num).ToString() + " " + index[loc];
            }
            catch { throw; }
        }

        /// <summary>
        /// Determines whether the folder has any child item (file or folder), per the
        /// File Manager contract. Uses the Graph-provided Folder.ChildCount metadata
        /// (zero extra API calls) and falls back to enumerating children only when
        /// the count is not available. A failed Graph request throws rather than
        /// being reported as an empty folder.
        /// </summary>
        private async Task<bool> HasChild(DriveItem item)
        {
            // Graph returns folder.ChildCount with the item metadata - no extra API call needed.
            if (item?.Folder?.ChildCount != null)
            {
                return item.Folder.ChildCount.Value > 0;
            }
            // Fallback: enumerate the first page only - HasChild just needs existence, not the full list.
            var children = await graphServiceClient.Drives[driveId].Items[item.Id].Children.GetAsync();
            return children?.Value?.Count > 0;
        }

        /// <summary>
        /// Optimized: Get parent ID path directly from ParentReference without recursive API calls
        /// </summary>
        private string GetFilterIdFromParentReference(DriveItem item)
        {
            // For now, return empty string since FilterId is mainly for internal tracking
            // and File Manager operations use actual ID field instead
            // This can be lazily computed if needed in future
            return "";
        }

        /// <summary>
        /// Optimized: Extract folder path directly from ParentReference.Path without recursive API calls
        /// Graph API provides: "/drives/{driveId}/root:/folder1/folder2"
        /// We extract: "/folder1/folder2/"
        /// </summary>
        private string GetFilterPathFromParentReference(DriveItem item)
        {
            if (item == null || item.ParentReference?.Path == null)
                return "";

            try
            {
                // Parse the path from Graph format: "/drives/{driveId}/root:/folder1/folder2"
                var pathParts = item.ParentReference.Path.Split(new[] { ":/" }, StringSplitOptions.RemoveEmptyEntries);
                if (pathParts.Length <= 1)
                    return "";

                var folderPath = pathParts[1];
                if (string.IsNullOrEmpty(folderPath))
                    return "/";

                // Ensure proper format: "/folder1/folder2/"
                var normalizedPath = "/" + folderPath.Trim('/') + "/";
                return normalizedPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] GetFilterPathFromParentReference parsing failed: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// DEPRECATED: Use GetFilterIdFromParentReference instead
        /// Kept for backward compatibility if needed
        /// </summary>
        private async Task<string> GetFilterId(DriveItem item)
        {
            if (item.ParentReference != null && item.ParentReference.Id != null)
            {
                var currentItem = item;
                var parentIds = new List<string>();
                while (currentItem?.ParentReference != null && currentItem.ParentReference.Id != null)
                {
                    parentIds.Insert(0, currentItem.ParentReference.Id);
                    
                    // Safety check: Ensure ParentReference.Id exists before fetching
                    if (string.IsNullOrEmpty(currentItem.ParentReference?.Id))
                    {
                        break;  // Stop traversing if we can't get parent ID
                    }
                    
                    try
                    {
                        currentItem = await graphServiceClient.Drives[driveId].Items[currentItem.ParentReference.Id].GetAsync();
                    }
                    catch
                    {
                        break;  // Stop on any API error and use what we have so far
                    }
                }
                return string.Join("/", parentIds) + "/";
            }
            return "";
        }

        /// <summary>
        /// DEPRECATED: Use GetFilterPathFromParentReference instead
        /// Kept for backward compatibility if needed
        /// </summary>
        private async Task<string> GetFilterPath(DriveItem item)
        {
            if (item.ParentReference != null && item.ParentReference.Path != null)
            {
                var currentItem = item;
                var parentPaths = new List<string>();
                while (currentItem?.ParentReference != null && currentItem.ParentReference.Path != null)
                {
                    var pathParts = currentItem.ParentReference.Path.Split(new[] { ":/" }, StringSplitOptions.RemoveEmptyEntries);
                    if (pathParts.Length > 1)
                    {
                        var folderNames = pathParts[1].Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                        if (folderNames.Any())
                        {
                            parentPaths.Insert(0, folderNames.Last());
                        }
                    }
                    
                    // Safety check: Ensure ParentReference.Id exists before fetching
                    if (string.IsNullOrEmpty(currentItem.ParentReference?.Id))
                    {
                        break;  // Stop traversing if we can't get parent ID
                    }
                    
                    try
                    {
                        currentItem = await graphServiceClient.Drives[driveId].Items[currentItem.ParentReference.Id].GetAsync();
                    }
                    catch
                    {
                        break;  // Stop on any API error and use what we have so far
                    }
                }
                var filterPath = "/" + string.Join("/", parentPaths);
                if (parentPaths.Count > 0) filterPath += "/";
                return filterPath;
            }
            return "";
        }

        protected virtual AccessPermission GetPathPermission(string path, bool isFile)
        {
            string[] fileDetails = GetFolderDetails(path);
            if (isFile)
            {
                return GetPermission(fileDetails[0].TrimStart('/') + fileDetails[1], fileDetails[1], true);
            }
            return GetPermission(fileDetails[0].TrimStart('/') + fileDetails[1], fileDetails[1], false);
        }

        protected virtual string[] GetFolderDetails(string path)
        {
            string[] str_array = path.Split('/'), fileDetails = new string[2];
            string parentPath = "";
            for (int i = 0; i < str_array.Length - 1; i++)
            {
                parentPath += str_array[i] + "/";
            }
            fileDetails[0] = parentPath;
            fileDetails[1] = str_array[str_array.Length - 1];
            return fileDetails;
        }

        protected virtual AccessPermission GetPermission(string location, string name, bool isFile)
        {
            AccessPermission permission = new AccessPermission();
            if (!isFile)
            {
                if (this.AccessDetails.AccessRules == null) { return null; }
                foreach (AccessRule folderRule in AccessDetails.AccessRules)
                {
                    if (folderRule.Path != null && folderRule.IsFile == false && (folderRule.Role == null || folderRule.Role == AccessDetails.Role))
                    {
                        if (folderRule.Path.IndexOf("*") > -1)
                        {
                            string parentPath = folderRule.Path.Substring(0, folderRule.Path.IndexOf("*"));
                            if ((location).IndexOf((parentPath)) == 0 || parentPath == "")
                            {
                                permission = UpdateFolderRules(permission, folderRule);
                            }
                        }
                        else if ((folderRule.Path) == (location) || (folderRule.Path) == (location + Path.DirectorySeparatorChar) || (folderRule.Path) == (location + "/"))
                        {
                            permission = UpdateFolderRules(permission, folderRule);
                        }
                        else if ((location).IndexOf((folderRule.Path)) == 0)
                        {
                            permission = UpdateFolderRules(permission, folderRule);
                        }
                    }
                }
                return permission;
            }
            else
            {
                if (this.AccessDetails.AccessRules == null) return null;
                string nameExtension = Path.GetExtension(name).ToLower();
                string fileName = Path.GetFileNameWithoutExtension(name);
                string currentPath = (location + "/");
                foreach (AccessRule fileRule in AccessDetails.AccessRules)
                {
                    if (!string.IsNullOrEmpty(fileRule.Path) && fileRule.IsFile && (fileRule.Role == null || fileRule.Role == AccessDetails.Role))
                    {
                        if (fileRule.Path.IndexOf("*.*") > -1)
                        {
                            string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf("*.*"));
                            if (currentPath.IndexOf((parentPath)) == 0 || parentPath == "")
                            {
                                permission = UpdateFileRules(permission, fileRule);
                            }
                        }
                        else if (fileRule.Path.IndexOf("*.") > -1)
                        {
                            string pathExtension = Path.GetExtension(fileRule.Path).ToLower();
                            string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf("*."));
                            if (((parentPath) == currentPath || parentPath == "") && nameExtension == pathExtension)
                            {
                                permission = UpdateFileRules(permission, fileRule);
                            }
                        }
                        else if (fileRule.Path.IndexOf(".*") > -1)
                        {
                            string pathName = Path.GetFileNameWithoutExtension(fileRule.Path);
                            string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf(pathName + ".*"));
                            if (((parentPath) == currentPath || parentPath == "") && fileName == pathName)
                            {
                                permission = UpdateFileRules(permission, fileRule);
                            }
                        }
                        else if ((fileRule.Path) == (Path.GetFileNameWithoutExtension(location)) || fileRule.Path == location || (fileRule.Path + nameExtension == location))
                        {
                            permission = UpdateFileRules(permission, fileRule);
                        }
                    }
                }
                return permission;
            }
        }

        protected virtual AccessPermission UpdateFolderRules(AccessPermission folderPermission, AccessRule folderRule)
        {
            folderPermission.Copy = HasPermission(folderRule.Copy);
            folderPermission.Download = HasPermission(folderRule.Download);
            folderPermission.Write = HasPermission(folderRule.Write);
            folderPermission.WriteContents = HasPermission(folderRule.WriteContents);
            folderPermission.Read = HasPermission(folderRule.Read);
            folderPermission.Upload = HasPermission(folderRule.Upload);
            folderPermission.Message = string.IsNullOrEmpty(folderRule.Message) ? string.Empty : folderRule.Message;
            return folderPermission;
        }

        protected virtual AccessPermission UpdateFileRules(AccessPermission filePermission, AccessRule fileRule)
        {
            filePermission.Copy = HasPermission(fileRule.Copy);
            filePermission.Download = HasPermission(fileRule.Download);
            filePermission.Write = HasPermission(fileRule.Write);
            filePermission.Read = HasPermission(fileRule.Read);
            filePermission.Message = string.IsNullOrEmpty(fileRule.Message) ? string.Empty : fileRule.Message;
            return filePermission;
        }

        protected virtual bool HasPermission(Syncfusion.Web.FileManager.Base.Permission rule)
        {
            return rule == Syncfusion.Web.FileManager.Base.Permission.Allow;
        }

        /// <summary>
        /// Determines if a string is a Microsoft Graph item ID (not a folder name/path)
        /// Graph IDs are typically 32+ alphanumeric characters (base64URL encoded)
        /// </summary>
        private bool IsGraphItemId(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            // Remove path delimiters
            var cleanValue = value.Trim('/').Trim();

            // Graph IDs are typically 32-48 characters and contain base64URL characters
            // (alphanumeric, hyphen, underscore, rarely slash or colon)
            if (cleanValue.Length < 24)
                return false;

            // Check if it looks like a Graph ID (mostly alphanumeric with possible -_= chars)
            return cleanValue.All(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == '=');
        }

        public string ToCamelCase(FileManagerResponse userData)
        {
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            };

            return JsonSerializer.Serialize(userData, options);
        }
    }
}
