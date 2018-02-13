using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.ServiceModel.Description;

namespace Handy.Crm.Powershell.Cmdlets.Connectivity
{
    public sealed class CrmServiceConfiguration
    {
        private readonly string XRMServicesEndpoint = "XRMServices/2011/Organization.svc";

        public ClientCredentials ClientCredentials { get; private set; }
        public IServiceConfiguration<IOrganizationService> ServiceConfiguration { get; private set; }

        public CrmServiceConfiguration(Uri organizationUri, string userName, string password)
        {
            var clientCredentials = new ClientCredentials();
            ClientCredentials.UserName.UserName = userName;
            ClientCredentials.UserName.Password = password;

            InitializeServiceConfiguration(organizationUri);
        }

        public CrmServiceConfiguration(Uri organizationUri, ClientCredentials clientCredentials)
        {
            ClientCredentials = clientCredentials;

            InitializeServiceConfiguration(organizationUri);
        }

        private void InitializeServiceConfiguration(Uri serviceUri)
        {
            if (!serviceUri.AbsolutePath.EndsWith(XRMServicesEndpoint))
            {
                serviceUri = new Uri(serviceUri, XRMServicesEndpoint);
            }

            ServiceConfiguration = ServiceConfigurationFactory.CreateConfiguration<IOrganizationService>(serviceUri, true, null);
        }
    }
}
