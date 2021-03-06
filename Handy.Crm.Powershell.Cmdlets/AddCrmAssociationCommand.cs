﻿using System;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsCommon.Add, "CRMAssociation")]
    public class AddCrmAssociationCommand : CrmCmdletBase
    {
        [Parameter(
            Mandatory = true)]
        public string EntityName { get; set; }

        [Parameter(
            Mandatory = true)]
        public Guid Id { get; set; }

        [Parameter(
            Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string Relationship { get; set; }

        [Parameter(
            Mandatory = true)]
        public EntityReference[] RelatedEntity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            EntityReferenceCollection entityReferenceCollection = new EntityReferenceCollection();
            entityReferenceCollection.AddRange(RelatedEntity);

            Connection.Associate(EntityName, Id, new Relationship(Relationship), entityReferenceCollection);
        }
    }
}
