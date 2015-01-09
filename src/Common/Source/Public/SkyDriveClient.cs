using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Live
{
    public class SkyDriveClient
    {
        private SkyDriveFolderInfo _rootFolder;

        public abstract class SkyDriveItemInfo
        {
            protected SkyDriveClient Client { get; private set; }
            protected IDictionary<string, object> ItemInfo { get; private set; }
            public string Id { get; private set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string ParentId { get; private set; }
            public int Size { get; private set; }
            public string Link { get; private set; }
            public DateTime CreatedTime { get; private set; }

            protected SkyDriveItemInfo(SkyDriveClient owner, dynamic itemInfo)
            {
                if (owner == null) throw new ArgumentNullException("owner");
                Client = owner;
                AssignFrom(itemInfo as IDictionary<string, object>);
            }

            protected virtual void AssignFrom(IDictionary<string, object> source)
            {
                if (source != null)
                {
                    ItemInfo = source;
                    Id = source.Get("id");
                    Name = source.Get("name");
                    Description = source.Get("description");
                    ParentId = source.Get("parent_id");
                    Size = source.Get("size", 0);
                    Link = source.Get("link");
                    var date = source.Get("created_time") ?? "";
                    if (date != "")
                        CreatedTime = DateTime.Parse(date);
                }
            }

            public Task WritePropertiesAsync()
            {
                return Client.SetPropertiesAsync(Id, Name, Description);
            }

            public Task UpdatePropertiesAsync()
            {
                return Client.Client.GetAsync(Id).ContinueWith(r => AssignFrom(r.Result.Result), TaskContinuationOptions.ExecuteSynchronously);
            }

            public Task DeleteAsync()
            {
                return Client.DeleteAsync(Id);
            }

            public Task<string> GetSharedReadLinkAsync()
            {
                return Client.GetSharedReadLinkAsync(Id);
            }

            public Task<string> GetSharedEditLinkAsync()
            {
                return Client.GetSharedEditLinkAsync(Id);
            }
            public Task<string> GetEmbedLinkAsync()
            {
                return Client.GetEmbedLinkAsync(Id);
            }

        }
        
        public class SkyDriveFolderInfo : SkyDriveItemInfo
        {
            public int Count { get; private set; }
            public SkyDriveFolderInfo(SkyDriveClient owner, dynamic folderInfo) : base(owner, (object)folderInfo)
            {
            }

            protected override void AssignFrom(IDictionary<string, object> source)
            {
                base.AssignFrom(source);
                Count = source.Get("count", 0);
            }

            public Task<IEnumerable<SkyDriveItemInfo>> GetItemsAsync(bool includeFolders = true)
            {
                return Client.GetItemsAsync(Id, includeFolders);
            }

            public Task<IEnumerable<SkyDriveFolderInfo>> GetFoldersAsync()
            {
                return Client.GetFoldersAsync(Id);
            }

            public Task<IEnumerable<SkyDriveFileInfo>> GetFilesAsync()
            {
                return Client.GetFilesAsync(Id);
            }

            public Task<SkyDriveFileInfo> UploadFileAsync(Stream fileStream, string fileName, string parentFolderId = null)
            {
                return Client.UploadFileAsync(fileStream, fileName, Id);
            }

            public Task<SkyDriveFileInfo> UploadFileAsync(string fullFileName, string parentFolderId = null)
            {
                return Client.UploadFileAsync(fullFileName, Id);
            }
        }

        public class SkyDriveFileInfo : SkyDriveItemInfo
        {
            public string Source { get; private set; }

            public SkyDriveFileInfo(SkyDriveClient owner, dynamic fileInfo) : base(owner, (object)fileInfo)
            {
            }
            protected override void AssignFrom(IDictionary<string, object> source)
            {
                base.AssignFrom(source);
                Source = source.Get("source");
            }


            public Task DownloadFileAsync(string fullFileName)
            {
                return Client.DownloadFileAsync(fullFileName, Id);
            }
            public Task<Stream> DownloadFileAsync()
            {
                return Client.DownloadFileAsync(Id);
            }
        }




        public LiveConnectClient Client { get; private set; }

        public SkyDriveFolderInfo RootFolder
        {
            get
            {
                if (_rootFolder == null)
                {
                    Client.GetAsync("me/skydrive").ContinueWith(r => _rootFolder = new SkyDriveFolderInfo(this, r.Result.Result), TaskContinuationOptions.ExecuteSynchronously).Wait();
                }
                return _rootFolder;
            }
        }

        public SkyDriveClient(LiveConnectClient liveConnectClient)
        {
            if (liveConnectClient == null) throw new ArgumentNullException("liveConnectClient");
            Client = liveConnectClient;
        }


        protected SkyDriveItemInfo Create(dynamic itemInfo)
        {
            if (itemInfo == null) throw new ArgumentNullException("itemInfo");
            if (itemInfo.type == "file")
                return new SkyDriveFileInfo(this, itemInfo);
            if (itemInfo.type == "folder")
                return new SkyDriveFolderInfo(this, itemInfo);
            throw new NotImplementedException(itemInfo.type + ": Unknown item type.");
        }

        public SkyDriveItemInfo FindByName(string frientlyNamePath)
        {
            // ToDo: implement later
            throw new NotImplementedException();
        }

        public Task<IEnumerable<SkyDriveItemInfo>> GetItemsAsync(string parentFolderId = null, bool includeFolders = true)
        {
            var request = parentFolderId == null ? "me/skydrive/files" : (parentFolderId + "/files");
            if (includeFolders)
                request += "?filter=folders";

            return Client.GetAsync(request).ContinueWith(r =>
            {
                var iEnum = r.Result.Result.Values.GetEnumerator();
                iEnum.MoveNext();
                return ((IEnumerable)iEnum.Current).OfType<dynamic>().Select(v => (SkyDriveItemInfo)Create(v));
            }, TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task<IEnumerable<SkyDriveFolderInfo>> GetFoldersAsync(string parentFolderId = null)
        {
            return GetItemsAsync(parentFolderId).ContinueWith(r => r.Result.OfType<SkyDriveFolderInfo>(), TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task<IEnumerable<SkyDriveFileInfo>> GetFilesAsync(string parentFolderId = null)
        {
            return GetItemsAsync(parentFolderId, false).ContinueWith(r => r.Result.OfType<SkyDriveFileInfo>(), TaskContinuationOptions.ExecuteSynchronously);
        }


        public Task<SkyDriveFileInfo> UploadFileAsync(Stream fileStream, string fileName, string parentFolderId = null)
        {
            parentFolderId = parentFolderId ?? RootFolder.Id;
            return Client.UploadAsync(parentFolderId, fileName, fileStream, OverwriteOption.DoNotOverwrite)
                .ContinueWith(r => new SkyDriveFileInfo(this, r.Result.Result), TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task<SkyDriveFileInfo> UploadFileAsync(string fullFileName, string parentFolderId = null)
        {
            var stream = new FileStream(fullFileName, FileMode.Open, FileAccess.Read);
            return UploadFileAsync(stream, Path.GetFileName(fullFileName), parentFolderId).ContinueWith(r =>
            {
                stream.Close();
                return r.Result;
            }, TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task DownloadFileAsync(string fullFileName, string fileId)
        {
            return DownloadFileAsync(fileId).ContinueWith(r =>
            {
                if (r.Result != null)
                {
                    using (r.Result)
                    {
                        var stream = new FileStream(fullFileName, FileMode.Create, FileAccess.Write);
                        return r.Result.CopyToAsync(stream).ContinueWith(r1 => stream.Close(), TaskContinuationOptions.ExecuteSynchronously);
                    }
                }
                return Task.FromResult(0);
            }, TaskContinuationOptions.ExecuteSynchronously);
        }
        public Task<Stream> DownloadFileAsync(string fileId)
        {
            return Client.DownloadAsync(fileId + "/content").ContinueWith(r => r.Result.Stream, TaskContinuationOptions.ExecuteSynchronously);
        }


        public Task<SkyDriveItemInfo> GetPropertiesAsync(string itemId)
        {
            return Client.GetAsync(itemId).ContinueWith(r => Create(r.Result.Result), TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task<LiveOperationResult> SetPropertiesAsync(string itemId, string name, string description = null)
        {
            var body = new Dictionary<string, object>();
            if (name != null)
                body["name"] = name;
            if (description != null)
                body["description"] = description == "" ? null : description;
            return Client.PutAsync(itemId, body);
        }
        public Task<LiveOperationResult> DeleteAsync(string itemId)
        {
            return Client.DeleteAsync(itemId);
        }

        public Task<string> GetSharedReadLinkAsync(string itemId)
        {
            return Client.GetAsync(itemId + "/shared_read_link").ContinueWith(r => r.Result.Result.Get("link"), TaskContinuationOptions.ExecuteSynchronously);
        }

        public Task<string> GetSharedEditLinkAsync(string itemId)
        {
            return Client.GetAsync(itemId + "/shared_edit_link").ContinueWith(r => r.Result.Result.Get("link"), TaskContinuationOptions.ExecuteSynchronously);
        }
        public Task<string> GetEmbedLinkAsync(string itemId)
        {
            return Client.GetAsync(itemId + "/embed").ContinueWith(r => r.Result.Result.Get("embed_html"), TaskContinuationOptions.ExecuteSynchronously);
        }

    }
}