using Sharepoint.Data.Connector.Models;
using Sharepoint.Data.Connector.Persistence;
using Sharepoint.Data.Connector.Business.Queries;
using Sharepoint.Data.Connector.Business.Commands;

namespace Sharepoint.Data.Connector
{
    /// <summary>
    /// This class contains the implementation of methods to be used in Sharepoint Data Connector.
    /// </summary>
    public class SharepointDataContext : ISharepointDataContext
    {
        private readonly SharepointDataQueries _queries;

        private readonly SharepointDataCommands _commands;

        public SharepointDataContext(SharepointContextConfiguration configuration)
        {
            _queries = new SharepointDataQueries(configuration);
            _commands = new SharepointDataCommands(configuration);
        }

        /// <summary>
        /// Function to validate if exists a folder in Sharepoint site.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Folder existence.</returns>
        public async Task<bool> ExistsFolderAsync(string serverRelativeUrl) => await _queries.ExistsFolderAsync(serverRelativeUrl);

        /// <summary>
        /// Function to validate if exists a file in Sharepoint site by name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to validate.</param>
        /// <returns>File existence.</returns>
        public async Task<bool> ExistsFileAsync(string serverRelativeUrl, string fileName) => await _queries.ExistsFileAsync(serverRelativeUrl, fileName);

        /// <summary>
        /// Function to delete a folder from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Deleted folder from Sharepoint.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<bool> DeleteResourceAsync(string serverRelativeUrl) => await _commands.DeleteResourceAsync(serverRelativeUrl);

        /// <summary>
        /// Function to delete a file from a specific path and file name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Deleted file from Sharepoint.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<bool> DeleteFileAsync(string serverRelativeUrl, string fileName) => await _commands.DeleteFileAsync(serverRelativeUrl, fileName);

        /// <summary>
        /// Function to download a file from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Content file.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<byte[]> DownloadFileAsync(string serverRelativeUrl, string fileName) => await _queries.DownloadFileAsync(serverRelativeUrl, fileName);

        /// <summary>
        /// Function to create a folder to main server relative url.
        /// </summary>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFolder> CreateFolderAsync(string folderName) => await _commands.CreateFolderAsync(folderName);

        /// <summary>
        /// Function to create a folder for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFolder> CreateFolderAsync(string serverRelativeUrl, string folderName) => await _commands.CreateFolderAsync(serverRelativeUrl, folderName);

        /// <summary>
        /// Function to upload a file for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <param name="content">Content file.</param>
        /// <param name="overrride">Override file or not.</param>
        /// <returns>Sharepoint file information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFile> UploadFileAsync(string serverRelativeUrl, string fileName, byte[] content) => await _commands.UploadFileAsync(serverRelativeUrl, fileName, content);
    }
}