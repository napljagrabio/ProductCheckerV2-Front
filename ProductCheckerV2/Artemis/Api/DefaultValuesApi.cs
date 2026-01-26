using ProductCheckerV2.Artemis.Api.Response;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;

namespace ProductCheckerV2.Artemis.Api
{
    internal class DefaultValuesApi
    {
        private HttpClient _httpClient { get; set; }

        public DefaultValuesApi(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<DefaultValuesResponse> Get(string targetTable, string targetColumn, object postBody)
        {
            try
            {
                var response =  await _httpClient.PostAsJsonAsync($"api/default-values/{targetTable}/{targetColumn}/get-value", postBody);
                return await response.Content.ReadFromJsonAsync<DefaultValuesResponse>(new JsonSerializerOptions() {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch (Exception)
            {
                return new DefaultValuesResponse() {
                    Value = null
                };
            }
        }
    }
}
