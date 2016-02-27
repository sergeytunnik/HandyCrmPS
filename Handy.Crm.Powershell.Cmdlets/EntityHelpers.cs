using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace Handy.Crm.Powershell.Cmdlets
{
    static class EntityHelpers
    {
        public static void UnwrapAttributes(this Entity entity)
        {
            AttributeCollection attributes = entity.Attributes;

            foreach (var attribute in attributes)
            {
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
