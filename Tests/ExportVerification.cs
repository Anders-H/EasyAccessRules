using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using EasyAccessRules;

internal static class ExportVerification
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string P = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static XmlNamespaceManager Namespaces(XmlDocument document)
    {
        var ns = new XmlNamespaceManager(document.NameTable);
        ns.AddNamespace("w", W);
        ns.AddNamespace("pkg", P);
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        return ns;
    }

    private static void Check(bool value, string message)
    {
        if (!value) throw new Exception(message);
    }

    private static string Text(XmlNode node, XmlNamespaceManager ns)
        => string.Join("|", node.SelectNodes(".//w:t", ns).Cast<XmlNode>().Select(n => n.InnerText));

    private static string Hash(XmlNode node)
    {
        using (var sha = SHA256.Create())
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(node.OuterXml)));
    }

    [STAThread]
    private static void Main(string[] args)
    {
        Synthetic();
        RealDocument(args[0], args[1]);
        Console.WriteLine("PASS: all export checks.");
    }

    private static void Synthetic()
    {
        var source = new XmlDocument();
        source.LoadXml("<w:document xmlns:w='" + W + "' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><w:body>"
            + "<w:p><w:r><w:t>Excluded</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:headerReference w:type='default' r:id='h1'/><w:footerReference w:type='default' r:id='f1'/><w:pgSz w:w='100'/></w:sectPr></w:pPr></w:p>"
            + "<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:bookmarkStart w:id='1' w:name='lost'/><w:r><w:fldChar w:fldCharType='begin'/><w:instrText> REF lost </w:instrText><w:fldChar w:fldCharType='separate'/><w:t>Keep A</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "<w:p><w:bookmarkEnd w:id='1'/><w:r><w:fldChar w:fldCharType='end'/><w:t>Excluded result</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:pgSz w:orient='landscape'/></w:sectPr></w:pPr></w:p>"
            + "<w:p><w:bookmarkStart w:id='2' w:name='kept'/><w:fldSimple w:instr='REF missing'><w:r><w:t>Keep B</w:t></w:r></w:fldSimple><w:bookmarkEnd w:id='2'/></w:p>"
            + "<w:sectPr><w:pgSz w:orient='portrait'/></w:sectPr></w:body></w:document>");
        var ns = Namespaces(source);
        var a = source.SelectSingleNode("//w:p[w:r/w:t='Keep A']", ns);
        var b = source.SelectSingleNode("//w:p[w:fldSimple]", ns);
        var original = source.OuterXml;
        var result = WordXmlExporter.Create(source, new HashSet<XmlNode> { a, b });
        var body = result.SelectSingleNode("//w:body", ns);
        Check(Text(body, ns) == "Keep A|Keep B", "Only selected content, in source order");
        Check(body.LastChild.LocalName == "sectPr", "Final section must be last");
        Check(body.SelectNodes(".//w:sectPr", ns).Count == 2, "Excluded sections must not create pages");
        Check(body.SelectSingleNode("w:p/w:pPr/w:sectPr/w:pgSz/@w:orient", ns).Value == "landscape", "Landscape layout lost");
        Check(body.SelectNodes(".//w:sectPr/w:headerReference[@r:id='h1']", ns).Count == 2, "Inherited header missing");
        Check(body.SelectNodes(".//w:sectPr/w:footerReference[@r:id='f1']", ns).Count == 2, "Inherited footer missing");
        Check(body.SelectNodes(".//w:fldChar | .//w:instrText | .//w:fldSimple", ns).Count == 0, "Partial fields not flattened");
        Check(body.SelectNodes(".//w:bookmarkStart | .//w:bookmarkEnd", ns).Count == 2, "Bookmark cleanup incorrect");
        Check(source.OuterXml == original, "Source document changed");
        var firstOnly = WordXmlExporter.Create(source, new HashSet<XmlNode> { a });
        Check(firstOnly.SelectSingleNode("//w:body/w:sectPr/w:pgSz/@w:orient", ns).Value == "landscape", "Last selected section layout lost");
        Check(firstOnly.SelectNodes("//w:body/w:p", ns).Count == 1, "Empty trailing section remains");
        var secondOnly = WordXmlExporter.Create(source, new HashSet<XmlNode> { b });
        Check(Text(secondOnly, ns) == "Keep B", "Repeated export reused old selection");
        try { WordXmlExporter.Create(source, new HashSet<XmlNode>()); throw new Exception("Empty selection accepted"); }
        catch (InvalidOperationException) { }
        try { WordXmlExporter.Create(source, new HashSet<XmlNode> { source.CreateElement("p") }); throw new Exception("Foreign selection accepted"); }
        catch (InvalidOperationException) { }
        Console.WriteLine("PASS: synthetic selection, sections, inheritance, fields, bookmarks, repeat and invalid selections.");
    }

    private static IEnumerable<TreeNode> Nodes(TreeNodeCollection nodes)
    {
        foreach (TreeNode node in nodes)
        {
            yield return node;
            foreach (var child in Nodes(node.Nodes)) yield return child;
        }
    }

    private static void RealDocument(string path, string output)
    {
        var source = new XmlDocument();
        source.Load(path);
        var ns = Namespaces(source);
        var originalHash = Hash(source);
        var originalBody = source.SelectSingleNode("//w:document/w:body", ns);
        var blocks = MainWindow.EnumerateBlocks(originalBody).ToList();
        var flags = BindingFlags.Instance | BindingFlags.NonPublic;
        using (var form = new MainWindow())
        {
            typeof(MainWindow).GetProperty("Document", flags).SetValue(form, source);
            typeof(MainWindow).GetMethod("BuildTree", flags).Invoke(form, null);
            var tree = (TreeView)typeof(MainWindow).GetField("treeView1", flags).GetValue(form);
            var nodes = Nodes(tree.Nodes).ToList();
            var article = nodes.Single(n => n.Text == "Article 1 - Subject matter and scope");
            article.Checked = true;
            var excluded = article.Nodes[1];
            excluded.Checked = false;
            var image = nodes.First(n => n.Text == "[Image / object]");
            var table = nodes.First(n => ((XmlNode)n.Tag).LocalName == "tbl");
            image.Checked = true;
            table.Checked = true;
            var selected = new HashSet<XmlNode>();
            typeof(MainWindow).GetMethod("CollectCheckedBlocks", BindingFlags.Static | BindingFlags.NonPublic)
                .Invoke(null, new object[] { tree.Nodes, selected });
            Check(!selected.Contains((XmlNode)excluded.Tag), "Unchecked child included by checked heading");
            var result = WordXmlExporter.Create(source, selected);
            var body = result.SelectSingleNode("//w:document/w:body", ns);
            var expected = string.Join("|", blocks.Where(selected.Contains).Select(b => Text(b, ns)).Where(t => t.Length > 0));
            Check(Text(body, ns) == expected, "Exported text differs from exact checked selection/order");
            Check(body.SelectNodes("w:tbl", ns).Count == 1, "Selected table missing");
            Check(body.SelectSingleNode(".//w:drawing | .//w:pict | .//w:object", ns) != null, "Selected image missing");

            var originalParts = source.SelectNodes("/pkg:package/pkg:part", ns).Cast<XmlElement>().ToList();
            var exportedParts = result.SelectNodes("/pkg:package/pkg:part", ns).Cast<XmlElement>().ToList();
            Check(originalParts.Count == exportedParts.Count, "Package parts missing");
            for (var i = 0; i < originalParts.Count; i++)
            {
                if (originalParts[i].GetAttribute("name", P) != "/word/document.xml")
                    Check(originalParts[i].OuterXml == exportedParts[i].OuterXml, "Package resource modified");
            }

            WordXmlExporter.Save(result, output);
            var reloaded = new XmlDocument();
            reloaded.Load(output);
            Check(Text(reloaded.SelectSingleNode("//w:document/w:body", ns), ns) == expected, "Save/load changed content");
            Check(reloaded.SelectSingleNode("processing-instruction('mso-application')") != null, "Word file association missing");
            WordXmlExporter.Save(result, output);
            Check(File.Exists(output), "Replacing existing export failed");
            Check(Hash(source) == originalHash, "Export changed source XML");
            Console.WriteLine("PASS: full XML input, checked chapter with excluded child, table/image, all " + originalParts.Count + " package parts, save/reload/replace, source unchanged.");
        }

        var all = WordXmlExporter.Create(source, new HashSet<XmlNode>(blocks));
        var allBody = all.SelectSingleNode("//w:document/w:body", ns);
        Check(Text(allBody, ns) == Text(originalBody, ns), "Full selection lost or duplicated text");
        Check(allBody.SelectNodes(".//w:sectPr", ns).Count == 26, "Full selection lost section properties");
        Check(allBody.SelectNodes(".//w:fldChar | .//w:instrText | .//w:fldSimple", ns).Count == 0, "Full selection left dangling fields");
        var starts = allBody.SelectNodes(".//w:bookmarkStart/@w:id", ns).Cast<XmlNode>().Select(n => n.Value).OrderBy(v => v);
        var ends = allBody.SelectNodes(".//w:bookmarkEnd/@w:id", ns).Cast<XmlNode>().Select(n => n.Value).OrderBy(v => v);
        Check(starts.SequenceEqual(ends), "Unpaired bookmarks remain");
        Console.WriteLine("PASS: full-document export, text preservation, 26 sections and all bookmark pairs.");
    }
}
