using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsLifecycle.Invoke, "CRMWhoAmI")]
    public class InvokeCrmWhoAmICommand : CrmCmdletBase
    {
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            var whoAmI = (WhoAmIResponse)Connection.Execute(
                new WhoAmIRequest());

            WriteObject(whoAmI);
        }
    }
}