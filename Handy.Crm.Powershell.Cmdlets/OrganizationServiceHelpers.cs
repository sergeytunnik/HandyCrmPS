using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
	static class OrganizationServiceHelpers
	{
		public static List<Entity> RetrieveMultipleAll(this OrganizationService organizationService, FetchExpression query)
		{
			int fetchCount = 5000;
			int pageNumber = 1;
			string pagingCookie = null;

			EntityCollection pageResult;
			List<Entity> result = new List<Entity>();

			do
			{
				query.MakePaged(fetchCount, pageNumber, pagingCookie);

				pageResult = organizationService.RetrieveMultiple(query);
				result.AddRange(pageResult.Entities.ToList());

				pageNumber++;
				pagingCookie = pageResult.PagingCookie;

			} while (pageResult.MoreRecords);

			return result;
		}
	}
}
