using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sharepoint.Http.Data.Connector.Models;
using Sharepoint.Http.Data.Connector.Extensions;
using Sharepoint.Http.Data.Connector.Business.Configurations;

namespace Sharepoint.Http.Data.Connector.Business.Queries
{
    /// <summary>
    /// This class contains all the operations that don't affect the information hosted in a Sharepoint site.
    /// </summary>
    public class SharepointDataQueries : SharepointDataConfiguration
    {
        public SharepointDataQueries(SharepointContextConfiguration configuration) : base(configuration) { }

        /// <summary>
        /// Fuction to retrive information of a resource that is in recycle bin using an unique identifier.
        /// </summary>
        /// <param name="resourceId">Resource unique identifier.</param>
        /// <returns>Recycle bin resource information.</returns>
        public async Task<SharepointRecycleResource?> GetRecycleBinResourceByIdAsync(Guid resourceId)
        {
            try
            {
                var client = await ConfigureClient(HeaderActionTypes.APPJSON_NOMETADATA);
                var responseHttp = await client.GetAsync($"_api/web/recyclebin('{ resourceId }')");
                if (!responseHttp.IsSuccessStatusCode)
                {
                    if (responseHttp.StatusCode == HttpStatusCode.NotFound)
                        return null;
                    await responseHttp.ValidateException();
                }
                return JsonConvert.DeserializeObject<SharepointRecycleResource>(await responseHttp.Content.ReadAsStringAsync());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// Function to validate if exists a folder in Sharepoint site.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Folder existence.</returns>
        public async Task<bool> ExistsFolderAsync(string serverRelativeUrl)
        {
            try
            {
                var client = await ConfigureClient(HeaderActionTypes.APPJSON_NOMETADATA);
                var responseHttp = await client.GetAsync($"_api/web/GetFolderByServerRelativeUrl('{_configuration.ServerRelativeUrl}{serverRelativeUrl}')/exists");
                if (!responseHttp.IsSuccessStatusCode)
                {
                    if(responseHttp.StatusCode == HttpStatusCode.NotFound)
                        return false;
                    await responseHttp.ValidateException();
                }    
                var response = JObject.Parse(await responseHttp.Content.ReadAsStringAsync());
                return response.Value<bool>("value");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// Function to validate if exists a file in Sharepoint site by name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to validate.</param>
        /// <returns>File existence.</returns>
        public async Task<bool> ExistsFileAsync(string serverRelativeUrl, string fileName)
        {
            try
            {
                var client = await ConfigureClient(HeaderActionTypes.APPJSON_NOMETADATA);
                var responseHttp = await client.GetAsync($"_api/web/GetFolderByServerRelativeUrl('{_configuration.ServerRelativeUrl}{serverRelativeUrl}')/files('{fileName}')/exists");
                if (!responseHttp.IsSuccessStatusCode)
                {
                    if (responseHttp.StatusCode == HttpStatusCode.NotFound)
                        return false;
                    await responseHttp.ValidateException();
                }
                var response = JObject.Parse(await responseHttp.Content.ReadAsStringAsync());
                return response.Value<bool>("value");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// Function to download a file from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Content file.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<byte[]?> DownloadFileAsync(string serverRelativeUrl, string fileName)
        {
            try
            {
                var client = await ConfigureClient(HeaderActionTypes.DOWNLOAD_FILE);
                var responseHttp = await client.GetAsync($"{_configuration.SharepointSiteUrl}_api/web/GetFolderByServerRelativeUrl('{_configuration.ServerRelativeUrl}{serverRelativeUrl}')/files('{fileName}')/$value");
                if (!responseHttp.IsSuccessStatusCode)
                {
                    if (responseHttp.StatusCode == HttpStatusCode.NotFound)
                        return null;
                    await responseHttp.ValidateException();
                }
                return await responseHttp.Content.ReadAsByteArrayAsync();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
    }
}
