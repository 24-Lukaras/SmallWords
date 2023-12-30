using System.Xml;

namespace SmallWords
{
    internal sealed class WRun
    {
        public string Text { get; set; }
        public List<WRunProperty> Properties { get; set; }
        public XmlNode Node { get; private set; }

        public int CollectionIndex { get; set; }

        public WRun() { }

        public WRun(XmlNode runNode)
        {
            var childNodes = runNode.ChildNodes.Cast<XmlNode>();
            var textNodes = childNodes.Where(x => x.Name == "w:t");
            var propertyNodes = childNodes.FirstOrDefault(x => x.Name == "w:rPr")?.ChildNodes;

            Node = runNode;
            Text = string.Join(string.Empty, textNodes.Select(x => x.InnerText));

            Properties = new List<WRunProperty>();
            if (propertyNodes != null)
            {
                foreach (XmlNode propertyNode in propertyNodes)
                {
                    Properties.Add(new WRunProperty(propertyNode));
                }
            }
        }

        public List<WRun> Split(params int[] indexes)
        {
            var orderedIndexes = indexes.OrderDescending();

            var result = new List<WRun>();

            string current = Text;
            foreach (int index in orderedIndexes)
            {
                string left = current.Substring(0, index);
                string right = current.Substring(index);
                var newNode = Node.Clone();
                var textNode = newNode.ChildNodes.Cast<XmlNode>().FirstOrDefault(x => x.Name == "w:t");
                textNode.InnerText = right;
                if (char.IsWhiteSpace(right[0]) || char.IsWhiteSpace(right[right.Length - 1]))
                {
                    var spaceAttribute = textNode.OwnerDocument.CreateAttribute("xml:space");
                    spaceAttribute.Value = "preserve";
                    textNode.Attributes.Append(spaceAttribute);
                }
                Node.ParentNode.InsertAfter(newNode, Node);
                var wRun = new WRun(newNode);
                wRun.CollectionIndex = CollectionIndex + index;
                result.Insert(0, wRun);
                current = left;
            }
            var currentTextNode = Node.ChildNodes.Cast<XmlNode>().FirstOrDefault(x => x.Name == "w:t");
            currentTextNode.InnerText = current;
            if (char.IsWhiteSpace(current[0]) || char.IsWhiteSpace(current[current.Length - 1]))
            {
                var spaceAttribute = Node.OwnerDocument.CreateAttribute("xml:space");
                spaceAttribute.Value = "preserve";
                currentTextNode.Attributes.Append(spaceAttribute);
            }
            Text = current;

            return result;
        }        
    }

    internal sealed class WRunCollection
    {
        public XmlNode Node { get; private set; }
        public IList<WRun> Collection { get; private set; }

        public WRunCollection(XmlNode node, IEnumerable<XmlNode> runNodes = null)
        {
            Node = node;

            var childNodes = node.ChildNodes.Cast<XmlNode>();
            if (runNodes == null)
            {
                runNodes = childNodes.Where(x => x.Name == "w:r"); ;
            }

            Collection = runNodes.Select(x => new WRun(x)).ToList();

            List<WRun> collection = new List<WRun>();
            int i = 0;
            foreach (XmlNode runNode in runNodes)
            {
                var wRun = new WRun(runNode);
                wRun.CollectionIndex = i;
                i += wRun.Text.Length;
                collection.Add(wRun);
            }
            Collection = collection;
        }

        public void Split(params int[] indexes)
        {
            Dictionary<int, List<WRun>> newRuns = new Dictionary<int, List<WRun>>();
            foreach (WRun wRun in Collection)
            {
                var runIndexes = indexes.Where(x => x > wRun.CollectionIndex && x < wRun.CollectionIndex + wRun.Text.Length).Select(x => x - wRun.CollectionIndex).ToArray();

                if (runIndexes.Any())
                {
                    var index = Collection.IndexOf(wRun);
                    newRuns.Add(index + 1, wRun.Split(runIndexes));
                }
            }

            foreach (var pair in newRuns)
            {
                foreach (var wRun in pair.Value)
                {
                    Collection.Insert(pair.Key, wRun);
                }
            }
        }

        public void MergeReplace(int start, int end, object value)
        {
            var mergeRuns = Collection.Where(x => x.CollectionIndex >= start && x.CollectionIndex + x.Text.Length <= end).OrderBy(x => x.CollectionIndex);

            string text = (value is string) ? value as string : value.ToString();

            var wRunsToDelete = mergeRuns.Skip(1).ToArray();
            var firstRun = mergeRuns.First();

            var textNode = firstRun.Node.ChildNodes.Cast<XmlNode>().FirstOrDefault(x => x.Name == "w:t");
            textNode.InnerText = text;
            firstRun.Text = text;

            foreach (var wRunToDelete in wRunsToDelete)
            {
                Collection.Remove(wRunToDelete);
                wRunToDelete.Node.ParentNode.RemoveChild(wRunToDelete.Node);
            }
        }


        public string GetCollectionString() => string.Join(string.Empty, Collection.Select(x => x.Text));
    }

    internal sealed class WRunProperty
    {
        public string Name { get; private set; }

        public string[] Attributes { get; private set; }


        public WRunProperty(XmlNode propertyNode)
        {
            Name = propertyNode.Name;
            Attributes = propertyNode.Attributes?.Cast<XmlAttribute>().Select(x => x.Value).ToArray() ?? Array.Empty<string>();
        }
    }
}
