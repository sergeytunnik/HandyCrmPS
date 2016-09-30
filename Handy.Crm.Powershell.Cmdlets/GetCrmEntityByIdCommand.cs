using System;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Handy.Crm.Powershell.Cmdlets
{
  [Cmdlet(VerbsCommon.Get, "CRMEntityById")]
  [OutputType("Microsoft.Xrm.Sdk.Entity")]
  public class GetCrmEntityByIdCommand : CrmCmdletBase
  {
    [Parameter(
      Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string EntityName { get; set; }

    [Parameter(
      Mandatory = true,
      ParameterSetName = "Columns")]
    [ValidateNotNullOrEmpty]
    public string[] Columns { get; set; }

    [Parameter(
      Mandatory = true,
      ParameterSetName = "AllColumns")]
    public SwitchParameter AllColumns { get; set; }

    [Parameter(
      Mandatory = true,
      ValueFromPipeline = true)]
    [ValidateNotNull]
    public Guid Id { get; set; }

    protected override void ProcessRecord()
    {
      base.ProcessRecord();

      ColumnSet columnSet = AllColumns.IsPresent ? new ColumnSet((bool)AllColumns) : new ColumnSet(Columns);

      Entity entity = organizationService.Retrieve(EntityName, Id, columnSet);

      WriteObject(entity);
    }
  }
}