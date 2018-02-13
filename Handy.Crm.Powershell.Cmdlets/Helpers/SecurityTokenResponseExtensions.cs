using Microsoft.Xrm.Sdk.Client;
using System;

namespace Handy.Crm.Powershell.Cmdlets.Helpers
{
    public static class SecurityTokenResponseExtensions
    {
        public static bool HasExpired(this SecurityTokenResponse securityTokenResponse)
        {
            return securityTokenResponse == null || DateTime.UtcNow >= securityTokenResponse.Response.Lifetime.Expires;
        }
    }
}
