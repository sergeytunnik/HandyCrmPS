using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Handy.Crm.Powershell.Cmdlets
{
	static class EntityHelpers
	{
		public static void UnwrapAttributes(this Entity entity)
		{
			AttributeCollection attributes = entity.Attributes;

			for (var i=attributes.Count - 1; i>-1;i--)
			{
				var attribute = attributes.ElementAt(i);
				var psobj = attribute.Value as System.Management.Automation.PSObject;
				if (psobj != null)
					attributes[attribute.Key] = psobj.BaseObject;
			}
		}

		public static object UnwrapPSObject(this object obj)
		{
			var psobj = obj as System.Management.Automation.PSObject;
			if (psobj != null)
				return psobj.BaseObject;
			else
				return obj;
		}
	}
}
