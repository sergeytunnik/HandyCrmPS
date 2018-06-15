using System;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;


namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsCommon.Remove, "CRMAssociation")]
    public class RemoveCrmAssociationCommand : CrmCmdletBase
    {
        [Parameter(
            Mandatory = true)]
        public string EntityName { get; set; }

        [Parameter(
            Mandatory = true)]
        public Guid Id { get; set; }

        [Parameter(
            Mandatory = true)]
        public string Relationship { get; set; }

        [Parameter(
            Mandatory = true)]
        public EntityReference[] RelatedEntity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            EntityReferenceCollection entityReferenceCollection = new EntityReferenceCollection();
            entityReferenceCollection.AddRange(RelatedEntity);

            Connection.Disassociate(EntityName, Id, new Relationship(Relationship), entityReferenceCollection);
        }
    }
}
