using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Graph;

namespace TeamsMeetingBookFunc.Authentication
{
    class AzureTokenServiceAuthProvider : IAuthenticationProvider
    {
        readonly AzureServiceTokenProvider azureServiceTokenProvider = new AzureServiceTokenProvider();
        private readonly string resourceUrl;

        public AzureTokenServiceAuthProvider(string resourceUrl)
        {
            if (string.IsNullOrWhiteSpace(resourceUrl))
            {
                throw new ArgumentException("Valid resource URL must be provided.", nameof(resourceUrl));
            }

            this.resourceUrl = resourceUrl;
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var accessToken = await azureServiceTokenProvider.GetAccessTokenAsync(resourceUrl).ConfigureAwait(false);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        }
    }
}
