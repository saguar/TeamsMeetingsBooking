using System;
using System.Collections.Generic;
using System.Text;

namespace TeamsMeetingBookFunc.Helpers
{
    internal class ConfigConstants
    {
        // When changing values here, please update the ARM template
        internal const string TenantIdCfg = "TenantID";
        internal const string ClientIdCfg = "ClientID";
        internal const string UserPasswordCfg = "UserPassword";
        internal const string UserEmailCfg = "UserEmail";
        internal const string DefaultMeetingNameCfg = "DefaultMeetingName";
        internal const string DefaultMeetingDurationMinsCfg = "DefaultMeetingDurationMins";
        internal const string AuthenticationModeCfg = "AuthenticationMode";
        internal const string ManagedIdentity = "managedIdentity";
        internal const string ClientSecret = "clientSecret";
        internal const string UsernamePassword = "usernamePassword";
    }
}
