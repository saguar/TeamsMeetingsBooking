﻿using System.Security;
using Microsoft.Graph;
using Microsoft.Graph.Auth;

namespace TeamsMeetingBookFunc.Helpers
{
    internal static class RequestHelpers
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope", Justification = "securePassword used outside this scope")]
        internal static T AddAuthenticationToRequest<T>(this T request, string username, string password) where T : IBaseRequest
        {
            var securePassword = new SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);  // keystroke by keystroke

            return request
                .WithScopes(new[] { "https://graph.microsoft.com/.default" })
                .WithUsernamePassword(username, securePassword);
        }
    }
}
