using RF.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace RF.Parser
{
    public class DotNetParser : BaseParser
    {
        public DotNetParser()
        {
            var doc = new XmlDocument();
            doc.LoadXml(Resources.DotNetRegex);

            var regNodes = doc.DocumentElement.SelectNodes("/root/regex/item[@value]");

            foreach (var node in regNodes)
            {
                _regexList.Add(new Regex((node as XmlNode).Attributes["value"].Value, RegexOptions.IgnoreCase));
            }
        }
    }
}
