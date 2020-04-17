using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using TeamsMeetingBookFunc.Authentication;
using TeamsMeetingBookFunc.Helpers;
using TeamsMeetingBookFunc.Services;

[assembly: FunctionsStartup(typeof(TeamsMeetingBookFunc.Startup))]
namespace TeamsMeetingBookFunc
{
    public class Startup : FunctionsStartup
    {
        const string baseUrl = "https://graph.microsoft.com/";

        public override void Configure(IFunctionsHostBuilder builder)
        {
            if (builder is null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            var services = builder.Services.BuildServiceProvider();
            IConfiguration configuration = services.GetRequiredService<IConfiguration>();

            Uri authority = new Uri($"https://login.microsoftonline.com/{configuration.GetConnectionStringOrSetting(ConfigConstants.TenantIdCfg)}");

            IAuthenticationProvider authenticationProvider = null;

            if (configuration.IsUsernamePasswordAuth())
            {
                var app = PublicClientApplicationBuilder.Create(configuration.GetClientId())
                    .WithAuthority(authority)
                    .Build();

                authenticationProvider = new UsernamePasswordProvider(app);
            }

            if (configuration.IsClientCredentialAuth())
            {
                var app = ConfidentialClientApplicationBuilder.Create(configuration.GetClientId())
                    .WithClientSecret(configuration.GetAccountPassword())
                    .WithAuthority(authority)
                    .Build();
                authenticationProvider  = new ClientCredentialProvider(app);
            }

            if (configuration.IsManagedIdentityAuth())
            {
                // no Azure Token Service-based provider available in SDK yet
                authenticationProvider = new AzureTokenServiceAuthProvider(baseUrl);
            }

            if(authenticationProvider == null)
            {
                throw new InvalidOperationException("Unknown AuthenticationMode - use 'usernamePassword', 'clientCredentials' or 'managedIdentity'");
            }

            builder.Services.AddSingleton(authenticationProvider);

            // need to use beta endpoint if creating an online meeting with service principal
            var versionedBaseUrl = baseUrl + (configuration.IsUsingServicePrincipal() ? "beta" : "v1.0");

            builder.Services.AddSingleton<IGraphServiceClient, GraphServiceClient>(f => new GraphServiceClient(versionedBaseUrl, authenticationProvider));
            
            builder.Services.AddSingleton<IBookingService, BookingService>();
        }
    }
}