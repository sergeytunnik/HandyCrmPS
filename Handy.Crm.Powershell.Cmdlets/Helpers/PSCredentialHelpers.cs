using System.Management.Automation;
using System.ServiceModel.Description;

namespace Handy.Crm.Powershell.Cmdlets.Helpers
{
    static class PSCredentialHelpers
    {
        public static ClientCredentials ToClientCredentials(this PSCredential psCredential)
        {
            var networkCredential = psCredential.GetNetworkCredential();
            var clientCredentials = new ClientCredentials();

            clientCredentials.UserName.UserName = string.IsNullOrWhiteSpace(networkCredential.Domain) ?
                networkCredential.UserName : string.Join("\\", networkCredential.Domain, networkCredential.UserName);
            clientCredentials.UserName.Password = networkCredential.Password;

            return clientCredentials;
        }
    }
}
