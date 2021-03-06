﻿using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Handy.Crm.Powershell.Cmdlets.Helpers;

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

        [Parameter(
            Mandatory = false)]
        public SwitchParameter RetrieveAll { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            List<Entity> result = new List<Entity>();

            if (RetrieveAll)
            {
                result = Connection.RetrieveMultipleAll(new FetchExpression(FetchXML));
            }
            else
            {
                EntityCollection e = Connection.RetrieveMultiple(new FetchExpression(FetchXML));
                result = e.Entities.ToList<Entity>();
            }

            WriteObject(result);
        }
    }
}
