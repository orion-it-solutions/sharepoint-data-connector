using Newtonsoft.Json.Linq;
using Sharepoint.Http.Data.Connector.Models;

namespace Sharepoint.Http.Data.Connector.Business.Configurations
{
    /// <summary>
    /// This class contains the configuration of initialization of services to connect with Sharepoint site.
    /// </summary>
    public class SharepointDataConfiguration
    {
        protected enum HeaderActionTypes
        {
            DOWNLOAD_FILE,
            DELETE_RESOURCE,
            APPJSON_NOMETADATA
        };

        protected readonly SharepointContextConfiguration _configuration;

        public SharepointDataConfiguration(SharepointContextConfiguration configuration) => _configuration = configuration;

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
        protected async Task<HttpClient> ConfigureClient()
        {
            var client = new HttpClient() { BaseAddress = new Uri($"{_configuration.SharepointSiteUrl}") };
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {await GetToken()}");
            return client;
        }

        /// <summary>
        /// Functino to configure the client HTTP to connect with Sharepoint site and specific ActionType.
        /// </summary>
        /// <returns>Http client connection.</returns>
        protected async Task<HttpClient> ConfigureClient(HeaderActionTypes action)
        {
            var client = new HttpClient() { BaseAddress = new Uri($"{_configuration.SharepointSiteUrl}") };
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {await GetToken()}");
            switch (action)
            {
                case HeaderActionTypes.DOWNLOAD_FILE:
                    client.DefaultRequestHeaders.Add("Accept", "application/octet-stream");
                    client.DefaultRequestHeaders.Add("binaryStringRequestBody", "true");
                    break;
                case HeaderActionTypes.DELETE_RESOURCE:
                    client.DefaultRequestHeaders.Add("X-RequestDigest", "SHAREPOINT_FORM_DIGEST");
                    client.DefaultRequestHeaders.Add("IF-MATCH", "*");
                    client.DefaultRequestHeaders.Add("X-HTTP-Method", "DELETE");
                    break;
                case HeaderActionTypes.APPJSON_NOMETADATA:
                    client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
                    break;
                default:
                    break;
            }
            return client;
        }
    }
}
