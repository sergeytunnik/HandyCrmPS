using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Handy.Crm.Powershell.Cmdlets
{
  [Cmdlet(VerbsCommon.Get, "CRMEntity")]
  public class GetCrmEntityCommand : CrmCmdletBase
  {
    [Parameter(
      Mandatory = true,
      ValueFromPipeline = true)]
    [ValidateNotNullOrEmpty]
    public string FetchXML { get; set; }

    protected override void ProcessRecord()
    {
      base.ProcessRecord();

      EntityCollection e = OrgService.RetrieveMultiple(new FetchExpression(FetchXML));
      List<Entity> list = e.Entities.ToList<Entity>();

      WriteObject(list);
    }
  }
}
