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

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

			WriteVerbose("Creating OrganizationService");
			organizationService = new OrganizationService(Connection ?? GetConnectionFromPSVariable(GlobalConnectionVariableName));
		}

		protected override void EndProcessing()
		{
			WriteVerbose("Disposing OrganizationService");
			organizationService.Dispose();

			base.EndProcessing();
		}

		private CrmConnection GetConnectionFromPSVariable(string VariableName)
		{
			WriteVerbose(string.Format("Getting connection from global variable '{0}'.", VariableName));

			var psobj = SessionState.PSVariable.GetValue(string.Format("global:{0}", VariableName), null);

			if (psobj == null)
				throw new PSArgumentNullException(VariableName, string.Format("Couldn't find global variable '{0}'.", VariableName));

			var connectionobj = psobj.UnwrapPSObject();

			if (!(connectionobj is Microsoft.Xrm.Client.CrmConnection))
				throw new PSArgumentException(string.Format("Variable '{0}' isn't an instance of Microsoft.Xrm.Client.CrmConnection.", VariableName), VariableName);

			return (CrmConnection)connectionobj;
		}
	}
}
