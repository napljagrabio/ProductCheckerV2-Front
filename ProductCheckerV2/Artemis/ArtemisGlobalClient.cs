using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ProductCheckerV2.Artemis.Api;
using ProductCheckerV2.Common;

namespace ProductCheckerV2.Artemis
{
    internal class ArtemisGlobalClient
    {
        private static readonly HttpClient _httpClient;
        private static ArtemisGlobalClient _instance { get; set; }

        public ValidateListingsApi ValidateListingsApi { get; private set; }

        static ArtemisGlobalClient()
        {
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(ConfigurationManager.GetArtemisApiBaseUrl())
            };
        }

        private ArtemisGlobalClient()
        {
            ValidateListingsApi = new ValidateListingsApi(_httpClient);
        }

        public static ArtemisGlobalClient Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ArtemisGlobalClient();

                return _instance;
            }
        }
    }
}
