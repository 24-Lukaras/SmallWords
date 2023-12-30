using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace SmallWords.Replacing
{
    public class PodjebnikTokenReplacer : IDataReplacer
    {
        public void Replace(Stream stream, Dictionary<string, object> data)
        {
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Update))
            {
                var document = zip.GetEntry("word/document.xml");
                using (var fs = document.Open())
                {
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(fs);

                    var textNodes = xmlDocument.GetElementsByTagName("w:r");

                    Dictionary<XmlNode, List<XmlNode>> groupedNodes = new Dictionary<XmlNode, List<XmlNode>>();
                    foreach (XmlNode node in textNodes)
                    {
                        List<XmlNode> coll;
                        if (!groupedNodes.TryGetValue(node.ParentNode, out coll))
                        {
                            coll = new List<XmlNode>();
                            groupedNodes[node.ParentNode] = coll;
                        }
                        coll.Add(node);
                    }

                    List<WRunCollection> wRunCollections = new List<WRunCollection>();
                    foreach (var pair in groupedNodes)
                    {
                        wRunCollections.Add(new WRunCollection(pair.Key, pair.Value));
                    }

                    Regex podjebnikCaseRegex = new Regex(@"(_[\w]+(_[\w]+)*)");
                    foreach (var wRunCollection in wRunCollections)
                    {
                        var matches = podjebnikCaseRegex.Matches(wRunCollection.GetCollectionString());
                        foreach (Match match in matches)
                        {
                            if (data.TryGetValue(match.Value.Substring(1), out var val))
                            {
                                wRunCollection.Split(match.Index, match.Index + match.Length);
                                wRunCollection.MergeReplace(match.Index, match.Index + match.Length, val);
                            }
                        }
                    }
                    fs.Seek(0, SeekOrigin.Begin);
                    xmlDocument.Save(fs);
                }
            }
            
        }
    }
}
