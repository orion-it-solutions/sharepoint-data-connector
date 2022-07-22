using Sharepoint.Http.Data.Connector.Models;

namespace Sharepoint.Http.Data.Connector.Persistence
{
    /// <summary>
    /// This interface contains all the main methods to be used in Sharepoint Data Connector.
    /// </summary>
    public interface ISharepointDataContext
    {
        /// <summary>
        /// Fuction to retrive information of a resource that is in recycle bin using an unique identifier.
        /// </summary>
        /// <param name="resourceId">Resource unique identifier.</param>
        /// <returns>Recycle bin resource information.</returns>
        Task<SharepointRecycleResource?> GetRecycleBinResourceByIdAsync(Guid resourceId);

        /// <summary>
        /// Function to validate if exists a folder in Sharepoint site.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Folder existence.</returns>
        Task<bool> ExistsFolderAsync(string serverRelativeUrl);

        /// <summary>
        /// Function to validate if exists a file in Sharepoint site by name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to validate.</param>
        /// <returns>File existence.</returns>
        Task<bool> ExistsFileAsync(string serverRelativeUrl, string fileName);

        /// <summary>
        /// Function to delete a resource from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Deleted resource from Sharepoint.</returns>
        Task<bool> DeleteResourceAsync(string serverRelativeUrl);

        /// <summary>
        /// Function to delete a file from a specific path and file name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Deleted file from Sharepoint.</returns>
        Task<bool> DeleteFileAsync(string serverRelativeUrl, string fileName);

        /// <summary>
        /// Function to download a file from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Byte array with file content.</returns>
        Task<byte[]?> DownloadFileAsync(string serverRelativeUrl, string fileName);

        /// <summary>
        /// Function to create a folder to main server relative url.
        /// </summary>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        Task<SharepointFolder?> CreateFolderAsync(string folderName);

        /// <summary>
        /// Function to create a folder for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        Task<SharepointFolder?> CreateFolderAsync(string serverRelativeUrl, string folderName);

        /// <summary>
        /// Function to upload a file for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <param name="content">Content file.</param>
        /// <returns>Sharepoint file information.</returns>
        Task<SharepointFile?> UploadFileAsync(string serverRelativeUrl, string fileName, byte[] content);

        /// <summary>
        /// Fuction to move a resource to recycle bin in a sharepoint site an unique identifier.
        /// </summary>
        /// <param name="serverRelativeUrl">Resource unique identifier.</param>
        /// <returns>Recycle bin resource information.</returns>
        Task<SharepointRecycleResource?> DeleteResourceToRecycleBinByIdAsync(string serverRelativeUrl);

        /// <summary>
        /// Fuction to restore a resource that is in recycle bin folder using an unique identifier.
        /// </summary>
        /// <param name="resourceId">Resource unique identifier.</param>
        /// <returns>Resource restored.</returns>
        Task<bool?> RestoreRecycleBinResourceByIdAsync(Guid resourceId);
    }
}
