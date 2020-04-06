using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using TeamsMeetingBookFunc.Helpers;
using TeamsMeetingBookFunc.Services;

[assembly: FunctionsStartup(typeof(TeamsMeetingBookFunc.Startup))]
namespace TeamsMeetingBookFunc
{
	public class Startup : FunctionsStartup
	{
		public override void Configure(IFunctionsHostBuilder builder)
		{
			var services = builder.Services.BuildServiceProvider();
			IConfiguration config = services.GetRequiredService<IConfiguration>();

			builder.Services.AddSingleton(s =>
			{
				Uri authority = new Uri($"https://login.microsoftonline.com/{config.GetValue<string>(ConfigConstants.TenantIdCfg)}");
				return PublicClientApplicationBuilder.Create(config.GetValue<string>(ConfigConstants.ClientIdCfg))
				.WithAuthority(authority)
				.Build();
			});
			builder.Services.AddSingleton<IAuthenticationProvider, UsernamePasswordProvider>(s =>
			{
				return new UsernamePasswordProvider(
					s.GetRequiredService<IPublicClientApplication>(),
					new string[] { "https://graph.microsoft.com/.default" }
					);
			});
			builder.Services.AddSingleton<IGraphServiceClient, GraphServiceClient>(s=> {
				return new GraphServiceClient(
					s.GetRequiredService<IAuthenticationProvider>()
				);
			});
			builder.Services.AddSingleton<IBookingService, BookingService>();

		}

	}
}
