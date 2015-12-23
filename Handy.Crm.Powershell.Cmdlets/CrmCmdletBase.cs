using System;
using System.Management.Automation;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;

namespace Handy.Crm.Powershell.Cmdlets
{
	public class CrmCmdletBase : PSCmdlet
	{
		private const string GlobalConnectionVariableName = "HandyCrmDefaultConnection";
		protected OrganizationService organizationService { get; set; }

		[Parameter(
			Mandatory = false,
			Position = 0)]
		[ValidateNotNull]
		public CrmConnection Connection { get; set; }

		[Parameter(
			Mandatory = false)]
		[ValidateNotNull]
		public Guid CallerId { get; set; }

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

			CrmConnection crmConnection = Connection ?? GetConnectionFromPSVariable(GlobalConnectionVariableName);

			if (CallerId != default(Guid))
				crmConnection.CallerId = CallerId;

			WriteVerbose("Creating OrganizationService");
			organizationService = new OrganizationService(crmConnection);
		}

		protected override void EndProcessing()
		{
			WriteVerbose("Disposing OrganizationService");
			organizationService.Dispose();

			base.EndProcessing();
		}

		private CrmConnection GetConnectionFromPSVariable(string variableName)
		{
			WriteVerbose(string.Format("Getting connection from global variable '{0}'.", variableName));

			var psobj = SessionState.PSVariable.GetValue(string.Format("global:{0}", variableName), null);

			if (psobj == null)
				throw new PSArgumentNullException(variableName, string.Format("Couldn't find global variable '{0}'.", variableName));

			var connectionobj = psobj.UnwrapPSObject();

			if (!(connectionobj is Microsoft.Xrm.Client.CrmConnection))
				throw new PSArgumentException(string.Format("Variable '{0}' isn't an instance of Microsoft.Xrm.Client.CrmConnection.", variableName), variableName);

			return (CrmConnection)connectionobj;
		}
	}
}
