using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ProductCheckerV2.Artemis.Api
{
    internal class ValidateListingsApi
    {
        private HttpClient _httpClient { get; set; }

        public ValidateListingsApi(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<HttpResponseMessage> Execute(string[] listings)
        {
            var payload = new
            {
                ids = listings
            };

            var httpRequestMessage = new HttpRequestMessage()
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri($"{_httpClient.BaseAddress.AbsoluteUri.TrimEnd('/')}/listing_id/validate"),
                Content = new StringContent(
                    JsonSerializer.Serialize(payload),
                    Encoding.UTF8,
                    "application/json")
            };
            return await _httpClient.SendAsync(httpRequestMessage);
        }
    }
}
