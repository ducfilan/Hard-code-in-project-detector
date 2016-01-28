using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using RF.Properties;
using System.Xml;
using System;

namespace RF.Parser
{
    public class AspParser : BaseParser
    {
        List<string> _treeAttList = new List<string>();

        public AspParser()
        {
            var doc = new XmlDocument();
            doc.LoadXml(Resources.AspRegex);

            var treeNodes = doc.DocumentElement.SelectNodes("/root/tree/item[@value]");

            foreach (var node in treeNodes)
            {
                _treeAttList.Add((node as XmlNode).Attributes["value"].Value);
            }

            var regNodes = doc.DocumentElement.SelectNodes("/root/regex/item[@value]");

            foreach (var node in regNodes)
            {
                _regexList.Add(new Regex((node as XmlNode).Attributes["value"].Value, RegexOptions.IgnoreCase));
            }
        }

        /// <summary>
        /// The find matching in web form.
        /// </summary>
        /// <param name="filePath">
        /// The file path.
        /// </param>
        /// <param name="resultList">
        /// The result list.
        /// </param>
        public override List<Item> GetHardCodeList(string filePath)
        {
            var resultList = new List<Item>();
            HtmlDocument doc = new HtmlWeb().Load(filePath);
            string content = System.IO.File.ReadAllText(filePath);

            foreach (var asp in _treeAttList)
            {
                var items = doc.DocumentNode.SelectNodes("//@" + asp);

                if (items == null) continue;

                foreach (var item in items)
                {
                    if (item.Name == "a")
                        continue;

                    var attribute = item.Attributes.FirstOrDefault(p => p.Name.ToLower() == asp.ToLower());

                    if (string.IsNullOrWhiteSpace(attribute.Value) ||
                     attribute.Value == "#" ||
                     attribute.Value.ToLower().Contains("javascript:"))
                    {
                        continue;
                    }

                    var suggest = string.IsNullOrEmpty(attribute.Value)
                                       ? ""
                                       : string.Format(_suggestionPattern["asp"], attribute.Value);

                    resultList.Add(new Item
                    {
                        ExtractedValue = attribute.Value,
                        Original = item.OuterHtml,
                        Suggestion = suggest,
                        Line = item.Line,
                        FilePath = filePath
                    });
                }
            }


            foreach (var rex in _regexList)
            {
                var match = rex.Match(content);
                while (match.Success)
                {
                    if (string.IsNullOrWhiteSpace(match.Groups["value"].Value) ||
                     match.Groups["value"].Value == "#" ||
                     match.Groups["value"].Value.ToLower().Contains("<%= httpcontext.current.request.url.absolutepath %>") ||
                     match.Groups["value"].Value.ToLower().Contains("<%= httpcontext.current.request.url %>") ||
                     match.Groups["value"].Value.ToLower().Contains("javascript:"))
                    {
                        match = match.NextMatch();
                        continue;
                    }

                    var suggestPtt = rex.ToString().Contains("font") ?
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
