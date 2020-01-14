using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Import, "CRMImportMapping")]
    public class ImportCrmImportMappingCommand : CrmCmdletBase
    {
        [Parameter(Mandatory = true, ParameterSetName = "File")]
        [ValidateNotNullOrEmpty]
        public string Path { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "String"),
        ValidateNotNullOrEmpty]
        public string MappingsXml { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter ReplaceIds { get; set; }

        protected override void ProcessRecord()
        {
            WriteVerbose(string.Format("Parameter set name: {0}", ParameterSetName));

            string mappings;

            switch (ParameterSetName)
            {
                case "File":
                    mappings = File.ReadAllText(GetAbsoluteFilePath(Path));
                    break;

                case "String":
                    mappings = MappingsXml;
                    break;

                default:
                    throw new PSNotSupportedException($"Parameter set '{ParameterSetName}' is not supported.");
            }

            var request = new ImportMappingsImportMapRequest()
            {
                MappingsXml = mappings,
                ReplaceIds = ReplaceIds
            };

            var response = (ImportMappingsImportMapResponse)Connection.Execute(request);

            WriteObject(response);
        }
    }
}
