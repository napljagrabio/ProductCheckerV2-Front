using ProductCheckerV2.Artemis.Api.Response;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;

namespace ProductCheckerV2.Artemis.Api
{
    internal class AuthenticationApi
    {
        private HttpClient _httpClient { get; set; }

        public AuthenticationApi(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<AuthenticationResponse> Login(string username, string password)
        {
            try
            {
                var authenticationResult =  await _httpClient.PostAsJsonAsync($"login", new {
                    user_id = username,
                    password
                });

                return await authenticationResult.Content.ReadFromJsonAsync<AuthenticationResponse>(new JsonSerializerOptions() {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
