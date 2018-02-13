using Handy.Crm.Powershell.Cmdlets.Helpers;
using Microsoft.Xrm.Sdk.Client;

namespace Handy.Crm.Powershell.Cmdlets.Connectivity
{
    public class CrmService : OrganizationServiceProxy
    {
        public CrmService(CrmServiceConfiguration configuration)
            : base(configuration.ServiceConfiguration, configuration.ClientCredentials) { }

        protected override void ValidateAuthentication()
        {
            if (this.SecurityTokenResponse.HasExpired())
            {
                this.Authenticate();
            }

            base.ValidateAuthentication();
        }
    }
}
