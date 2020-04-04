using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Security;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;

namespace TeamsMeetingBookFunc
{
    internal static class RequestHelpers
    {
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope", Justification = "securePassword is used by caller")]
        internal static T AddAuthenticationToRequest<T>(this T request, string username, string password) where T : IBaseRequest
        {
            var securePassword = new SecureString();
            // you should fetch the password from Azure Keyvault
            foreach (char c in password)
                securePassword.AppendChar(c);  // keystroke by keystroke

            return request
                .WithScopes(new[] { "https://graph.microsoft.com/.default" })
                .WithUsernamePassword(username, securePassword);
        }
    }
}
