using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Collections;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsLifecycle.Invoke, "CRMOrganizationRequest")]
	public class InvokeCrmOrganizationRequestCommand : CrmCmdletBase
	{
		[Parameter(
			Mandatory = true)]
		[ValidateNotNullOrEmpty]
		public string RequestName { get; set; }

		[Parameter(
			Mandatory = false)]
		[ValidateNotNull]
		public Hashtable Parameters { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			OrganizationRequest orgRequest = new OrganizationRequest(RequestName)
			{
				Parameters = new ParameterCollection()
			};

			if (MyInvocation.BoundParameters.ContainsKey("Parameters"))
			{
				foreach (string key in Parameters.Keys)
				{
					orgRequest.Parameters.Add(new KeyValuePair<string, object>(key, Parameters[key].UnwrapPSObject()));
				}
			}

			OrganizationResponse orgResponse = (OrganizationResponse)organizationService.Execute(orgRequest);

			WriteObject(orgResponse);
		}
	}
}
