using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ProductCheckerV2.Artemis.Api.Response
{
    internal class AuthenticationResponse
    {
        public Meta Meta { get; set; }
    }

    internal class Meta
    {
        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; }
    }
}
