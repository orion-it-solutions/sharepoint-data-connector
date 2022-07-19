using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sharepoint.Data.Connector.Models;
using Sharepoint.Data.Connector.Persistence;

namespace Sharepoint.Data.Connector
{
    /// <summary>
    /// This class contains the implementation of methods to be used in Sharepoint Data Connector.
    /// </summary>
    public class SharepointDataContext : ISharepointDataContext
    {
        private readonly SharepointContextConfiguration _configuration;

        public SharepointDataContext(SharepointContextConfiguration configuration)
        {
            _configuration = configuration;
        }

        /// <summary>
        /// Function to get an access token to Sharepoint site.
        /// </summary>
        /// <returns>Token value</returns>
        /// <exception cref="Exception">Obtention token error.</exception>
        private async Task<string> GetToken()
        {
            var client = new HttpClient() { BaseAddress = new Uri($"{_configuration.AuthenticationUrl}") };
            var content = new[]
            {
                new KeyValuePair<string, string>("resource", _configuration.Resource),
                new KeyValuePair<string, string>("client_id", _configuration.ClientId),
                new KeyValuePair<string, string>("client_secret", _configuration.ClientSecret),
                new KeyValuePair<string, string>("grant_type", _configuration.GrantType),
            };
            var responseHttp = await client.PostAsync("tokens/oAuth/2", new FormUrlEncodedContent(content));
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            var response = JObject.Parse(await responseHttp.Content.ReadAsStringAsync());
            return response.Value<string>("access_token") ?? string.Empty;
        }

        /// <summary>
        /// Functino to configure the client HTTP to connect with Sharepoint site.
        /// </summary>
        /// <returns>Http client connection.</returns>
        private async Task<HttpClient> ConfigureClient()
        {
            var client = new HttpClient() { BaseAddress = new Uri($"{_configuration.SharepointSiteUrl}") };
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer { await GetToken() }");
            return client;
        }

        /// <summary>
        /// Function to validate if exists a folder in Sharepoint site.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Folder existence.</returns>
        public async Task<bool> ExistsFolderAsync(string serverRelativeUrl)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
            var responseHttp = await client.GetAsync($"_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }')/exists");
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            var response = JObject.Parse(await responseHttp.Content.ReadAsStringAsync());
            return response.Value<bool>("value");
        }

        /// <summary>
        /// Function to validate if exists a file in Sharepoint site by name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to validate.</param>
        /// <returns>File existence.</returns>
        public async Task<bool> ExistsFileAsync(string serverRelativeUrl, string fileName)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
            var responseHttp = await client.GetAsync($"_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }')/files('{ fileName }')/exists");
            if (!responseHttp.IsSuccessStatusCode)
            {
                if (responseHttp.StatusCode == HttpStatusCode.NotFound)
                    return false;
                throw new Exception();
            }
            var response = JObject.Parse(await responseHttp.Content.ReadAsStringAsync());
            return response.Value<bool>("value");
        }

        /// <summary>
        /// Function to delete a folder from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <returns>Deleted folder from Sharepoint.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<bool> DeleteResourceAsync(string serverRelativeUrl)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("X-RequestDigest", "SHAREPOINT_FORM_DIGEST");
            client.DefaultRequestHeaders.Add("IF-MATCH", "*");
            client.DefaultRequestHeaders.Add("X-HTTP-Method", "DELETE");
            var responseHttp = await client.PostAsync($"_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }')", null);
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            return responseHttp.IsSuccessStatusCode;
        }

        /// <summary>
        /// Function to delete a file from a specific path and file name.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Deleted file from Sharepoint.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<bool> DeleteFileAsync(string serverRelativeUrl, string fileName)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("X-RequestDigest", "SHAREPOINT_FORM_DIGEST");
            client.DefaultRequestHeaders.Add("IF-MATCH", "*");
            client.DefaultRequestHeaders.Add("X-HTTP-Method", "DELETE");
            var responseHttp = await client.PostAsync($"_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }/{ fileName }')", null);
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            return responseHttp.IsSuccessStatusCode;
        }

        /// <summary>
        /// Function to download a file from a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <returns>Content file.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<byte[]> DownloadFileAsync(string serverRelativeUrl, string fileName)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("Authorization", $"Bearer { await GetToken() }");
                client.Headers.Add(HttpRequestHeader.Accept, "application/octet-stream");
                client.Headers.Add("binaryStringRequestBody", "true");
                var endpointUri = new Uri($"{ _configuration.SharepointSiteUrl }_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }')/files('{ fileName }')/$value");
                var result = client.DownloadData(endpointUri);
                return result;
            }
        }

        /// <summary>
        /// Function to create a folder to main server relative url.
        /// </summary>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFolder> CreateFolderAsync(string folderName)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "_api/web/folders");
            request.Content = new StringContent(JsonConvert.SerializeObject(new { ServerRelativeUrl = $"{ _configuration.ServerRelativeUrl }{ folderName }" }), Encoding.UTF8, "application/json");
            var responseHttp = await client.SendAsync(request);
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            var response = JsonConvert.DeserializeObject<SharepointFolder>(await responseHttp.Content.ReadAsStringAsync());
            return response;
        }

        /// <summary>
        /// Function to create a folder for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="folderName">Folder name to be created.</param>
        /// <returns>Sharepoint folder information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFolder> CreateFolderAsync(string serverRelativeUrl, string folderName)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "_api/web/folders");
            request.Content = new StringContent(JsonConvert.SerializeObject(new { ServerRelativeUrl = $"{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }/{ folderName }" }), Encoding.UTF8, "application/json");
            var responseHttp = await client.SendAsync(request);
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            var response = JsonConvert.DeserializeObject<SharepointFolder>(await responseHttp.Content.ReadAsStringAsync());
            return response;
        }

        /// <summary>
        /// Function to upload a file for a specific path.
        /// </summary>
        /// <param name="serverRelativeUrl">Relative url of resource.</param>
        /// <param name="fileName">File name to delete.</param>
        /// <param name="content">Content file.</param>
        /// <param name="overrride">Override file or not.</param>
        /// <returns>Sharepoint file information.</returns>
        /// <exception cref="Exception">Sharepoint connection error.</exception>
        public async Task<SharepointFile> UploadFileAsync(string serverRelativeUrl, string fileName, byte[] content)
        {
            var client = await ConfigureClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
            var responseHttp = await client.PostAsync(
                $"_api/web/GetFolderByServerRelativeUrl('{ _configuration.ServerRelativeUrl }{ serverRelativeUrl }')/Files/add(overwrite=true, url='{ fileName }')",
                new ByteArrayContent(content)
            );
            if (!responseHttp.IsSuccessStatusCode)
                throw new Exception();
            var response = JsonConvert.DeserializeObject<SharepointFile>(await responseHttp.Content.ReadAsStringAsync());
            return response;
        }
    }
}