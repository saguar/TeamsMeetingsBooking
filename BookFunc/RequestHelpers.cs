using System;
using System.Collections.Generic;
using System.Security;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;

namespace TeamsMeetingBookFunc
{
    internal static class RequestHelpers
    {
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
