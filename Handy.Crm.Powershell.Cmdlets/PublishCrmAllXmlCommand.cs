using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Publish, "CRMAllXml")]
    public class PublishCrmAllXmlCommand : CrmCmdletBase
    {
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            Connection.Execute(new PublishAllXmlRequest());
        }
    }
}
