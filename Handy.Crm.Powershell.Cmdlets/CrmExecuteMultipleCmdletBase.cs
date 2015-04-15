using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	public class CrmExecuteMultipleCmdletBase : CrmCmdletBase
	{
		protected ExecuteMultipleRequest _executeMultipleRequest;

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

			WriteVerbose("Creating empty ExecuteMultipleRequest");
			_executeMultipleRequest = new ExecuteMultipleRequest()
			{
				Settings = new ExecuteMultipleSettings()
				{
					ContinueOnError = false,
					ReturnResponses = true
				},
				Requests = new OrganizationRequestCollection()
			};
		}

		protected override void EndProcessing()
		{
			WriteVerbose("Executing ExecuteMultipleRequest");
			ExecuteMultipleResponse executeMultipleResponse = (ExecuteMultipleResponse)OrgService.Execute(_executeMultipleRequest);

			ProcessResponse(executeMultipleResponse);

			base.EndProcessing();
		}

		protected virtual void ProcessResponse(ExecuteMultipleResponse response)
		{
			WriteObject(response);
		}
	}
}
