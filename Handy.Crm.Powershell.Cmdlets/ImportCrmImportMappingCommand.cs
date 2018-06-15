using System;
using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Import, "CRMImportMapping")]
    public class ImportCrmImportMappingCommand : CrmCmdletBase
    {
        private string _path;

        [Parameter(
            Mandatory = true,
            ParameterSetName = "File"),
        ValidateNotNullOrEmpty]
        public string Path
        {
            get
            {
                return _path;
            }

            set
            {
                _path = System.IO.Path.IsPathRooted(value)
                    ? value
                    : System.IO.Path.GetFullPath(SessionState.Path.CurrentLocation.Path + value);
            }
        }

        [Parameter(
            Mandatory = true,
            ParameterSetName = "String"),
        ValidateNotNullOrEmpty]
        public string MappingsXml { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ReplaceIds { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            WriteVerbose(string.Format("Parameter set name: {0}", ParameterSetName));

            string mappings = String.Empty;

            switch (ParameterSetName)
            {
                case "File":
                    mappings = File.ReadAllText(Path);
                    break;

                case "String":
                    mappings = MappingsXml;
                    break;
            }

            ImportMappingsImportMapRequest request = new ImportMappingsImportMapRequest()
            {
                MappingsXml = mappings,
                ReplaceIds = ReplaceIds
            };

            ImportMappingsImportMapResponse response = (ImportMappingsImportMapResponse)Connection.Execute(request);

            WriteObject(response);
        }
    }
}
