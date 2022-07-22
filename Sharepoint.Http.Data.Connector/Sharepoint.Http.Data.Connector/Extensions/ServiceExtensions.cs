using Newtonsoft.Json.Linq;
using Sharepoint.Http.Data.Connector.Business.Infrastructure.Exceptions;

namespace Sharepoint.Http.Data.Connector.Extensions
{
    public static class ServiceExtensions
    {
        public static async Task ValidateException(this HttpResponseMessage httpResponse)
        {
            var status = httpResponse.StatusCode;
            string responseBody = await httpResponse.Content.ReadAsStringAsync();
            if (string.IsNullOrEmpty(responseBody))
                httpResponse.EnsureSuccessStatusCode();
            var response = JObject.Parse(responseBody);
            switch (status)
            {
                case System.Net.HttpStatusCode.NotFound:
                    if((JObject)response["error"] is not null)
                        throw new NotFoundException((string)response["error"]["message"]["value"]);
                    if ((JObject)response["odata.error"] is not null)
                        throw new NotFoundException((string)response["odata.error"]["message"]["value"]);
                    throw new NotFoundException($"The resource does not exist.");
                case System.Net.HttpStatusCode.BadRequest:
                    if(!string.IsNullOrEmpty((string)response["error_description"]))
                        throw new BadRequestException((string)response["error_description"]);
                    throw new BadRequestException();
                case System.Net.HttpStatusCode.Unauthorized:
                    if (!string.IsNullOrEmpty((string)response["error_description"]))
                        throw new UnauthorizedException((string)response["error_description"]);
                    httpResponse.EnsureSuccessStatusCode();
                    break;
                case System.Net.HttpStatusCode.InternalServerError:
                    if(!string.IsNullOrEmpty((string)response["error"]["message"]["value"]))
                        throw new InternalServerException((string)response["error"]["message"]["value"]);
                    httpResponse.EnsureSuccessStatusCode();
                    break;
                default:
                    throw new Exception();
            }
        }
    }
}
