using Microsoft.Xrm.Sdk;
using System.Linq;
using System.Management.Automation;

namespace Handy.Crm.Powershell.Cmdlets.Helpers
{
    static class EntityHelpers
    {
        public static void UnwrapAttributes(this Entity entity)
        {
            AttributeCollection attributes = entity.Attributes;

            for (var i = attributes.Count - 1; i > -1; i--)
            {
                var attribute = attributes.ElementAt(i);

                var psobj = attribute.Value as PSObject;

                if (psobj != null)
                    attributes[attribute.Key] = psobj.BaseObject;
            }
        }

        public static object UnwrapPSObject(this object obj)
        {
            var psobj = obj as PSObject;

            if (psobj != null)
                return psobj.BaseObject;
            else
                return obj;
        }
    }
}
