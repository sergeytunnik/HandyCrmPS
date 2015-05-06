using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System.Management.Automation;

namespace Handy.Crm.Powershell.Cmdlets
{
	public class CrmExecuteMultipleCmdletBase : CrmCmdletBase
	{
		protected ExecuteMultipleRequest _executeMultipleRequest;

		// More about what is inside of reponses collection depending on settings at
		// https://msdn.microsoft.com/en-us/library/jj863631%28v=crm.5%29.aspx

		[Parameter(
			Mandatory = false)]
		public SwitchParameter ContinueOnError { get; set; }

		[Parameter(
			Mandatory = false)]
		public SwitchParameter ReturnResponses { get; set; }

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

			WriteVerbose("Creating empty ExecuteMultipleRequest");
			_executeMultipleRequest = new ExecuteMultipleRequest()
			{
				Settings = new ExecuteMultipleSettings()
				{
					ContinueOnError = ContinueOnError,
					ReturnResponses = ReturnResponses
				},
				Requests = new OrganizationRequestCollection()
			};
		}

		protected override void EndProcessing()
		{
			WriteVerbose("Executing ExecuteMultipleRequest");
			ExecuteMultipleResponse executeMultipleResponse = (ExecuteMultipleResponse)organizationService.Execute(_executeMultipleRequest);

			ProcessResponse(executeMultipleResponse);

			base.EndProcessing();
		}

		protected virtual void ProcessResponse(ExecuteMultipleResponse response)
		{
			WriteObject(response);
		}
	}
}
