using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace TeamsMeetingBookFunc.Helpers
{
    internal static class ConfigHelpers
    {
        internal static bool IsUsingServicePrincipal(this IConfiguration configuration)
        {
            return IsClientCredentialAuth(configuration) || IsManagedIdentityAuth(configuration);
        }

        internal static bool IsManagedIdentityAuth(this IConfiguration configuration)
        {
            return string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), ConfigConstants.ManagedIdentity, StringComparison.InvariantCultureIgnoreCase);
        }

        internal static bool IsClientCredentialAuth(this IConfiguration configuration)
        {
            return string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), ConfigConstants.ClientSecret, StringComparison.InvariantCultureIgnoreCase);
        }

        internal static bool IsUsernamePasswordAuth(this IConfiguration configuration)
        {
            return string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), ConfigConstants.UsernamePassword, StringComparison.InvariantCultureIgnoreCase);
        }

        internal static string GetAccountEmail(this IConfiguration configuration)
        {
            return configuration[ConfigConstants.UserEmailCfg];
        }

        internal static string GetAccountPassword(this IConfiguration configuration)
        {
            return configuration.GetConnectionStringOrSetting(ConfigConstants.UserPasswordCfg);
        }
        internal static string GetClientId(this IConfiguration configuration)
        {
            return configuration.GetConnectionStringOrSetting(ConfigConstants.ClientIdCfg);
        }
    }
}
