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
    public abstract class BaseParser : IParser
    {
        protected List<Regex> _regexList = new List<Regex>();
        protected Dictionary<string, string> _suggestionPattern = new Dictionary<string, string>();

        protected BaseParser()
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(Resources.SuggestionPattern);

            var nodes = doc.SelectNodes("/root/item[@name and @value]");
            foreach (var n in nodes)
            {
                var node = n as XmlNode;
                _suggestionPattern[node.Attributes.GetNamedItem("name").Value] = node.Attributes.GetNamedItem("value").Value;
            }
        }

        public virtual List<Item> GetHardCodeList(string filePath)
        {
            if (_regexList == null)
                return null;

            var resultList = new List<Item>();
            string content = System.IO.File.ReadAllText(filePath);

            foreach (var rex in _regexList)
            {
                var match = rex.Match(content);
                while (match.Success)
                {
                    if (rex.GetGroupNames().Contains("value"))
                    {
                        if (string.IsNullOrWhiteSpace(match.Groups["value"].Value) ||
                            match.Groups["value"].Value == "#" ||
                            match.Groups["value"].Value.ToLower().Contains("javascript:"))
                        {
                            match = match.NextMatch();
                            continue;
                        }
                    }

                    var suggestPtt = rex.ToString().Contains("font-") ? 
                        _suggestionPattern["font"] : 
                        _suggestionPattern["common"];
                    var suggest = string.IsNullOrEmpty(match.Groups["value"].Value)
                                       ? ""
                                       : string.Format(suggestPtt, match.Groups["value"].Value);

                    var lCount = content.Substring(0, match.Index).Split(new[] { Environment.NewLine }, StringSplitOptions.None).Count();

                    resultList.Add(new Item
                    {
                        ExtractedValue = match.Groups["value"].Value,
                        Original = match.Value,
                        Suggestion = suggest,
                        Line = lCount,
                        FilePath = filePath
                    });

                    match = match.NextMatch();
                }
            }

            return resultList;
        }
    }
}
