using System;
using System.Xml;
using Microsoft.Xrm.Sdk.Query;

namespace Handy.Crm.Powershell.Cmdlets
{
  static class FetchExpressionHelpers
  {
    public static void MakePaged(this FetchExpression fetchExpression, int count, int page, string pagingCookie)
    {
      XmlDocument doc = new XmlDocument();
      doc.LoadXml(fetchExpression.Query);

      XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

      XmlAttribute countAttr = doc.CreateAttribute("count");
      countAttr.Value = count.ToString();
      attrs.Append(countAttr);

      XmlAttribute pageAttr = doc.CreateAttribute("page");
      pageAttr.Value = page.ToString();
      attrs.Append(pageAttr);

      if (pagingCookie != null)
      {
        XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
        pagingAttr.Value = pagingCookie;
        attrs.Append(pagingAttr);
      }

      fetchExpression.Query = doc.OuterXml;
    }
  }
}
