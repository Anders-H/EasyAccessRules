#nullable enable
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;

namespace EasyAccessRules;

public partial class MainWindow : Form
{
    private XmlDocument Document { get; set; }
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private bool _updatingChecks;

    public MainWindow()
    {
        InitializeComponent();
        Document = new XmlDocument();
        treeView1.CheckBoxes = true;
        treeView1.HideSelection = false;
        treeView1.AfterCheck += TreeViewAfterCheck;
        treeView1.AfterSelect += TreeViewAfterSelect;
    }

    private void openToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var x = new OpenFileDialog();
        x.Title = @"Open XML File";
        x.Filter = @"XML files (*.xml)|*.xml|All files (*.*)|*.*";

        if (x.ShowDialog() != DialogResult.OK)
            return;

        Cursor = Cursors.WaitCursor;
        Document = new XmlDocument();
        Document.Load(x.FileName);
        treeView1.Nodes.Clear();
        textBox1.Text = "";
        BuildTree();
        Cursor = Cursors.Default;
    }

    private void BuildTree()
    {
        treeView1.BeginUpdate();

        try
        {
            treeView1.Nodes.Clear();
            textBox1.Clear();
            var namespaces = new XmlNamespaceManager(Document.NameTable);
            namespaces.AddNamespace("w", WordNamespace);
            namespaces.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");

            // Read only the document body, leaving images, styles and relationships in Document.
            var body = Document.SelectSingleNode("/pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body", namespaces) ?? Document.SelectSingleNode("/w:document/w:body", namespaces);
            
            if (body == null)
            {
                textBox1.Text = @"No Word document body was found in the XML file.";
                return;
            }

            var styles = new Dictionary<string, XmlNode>(StringComparer.Ordinal);
            var xmlStyles = Document.SelectNodes("//w:styles/w:style", namespaces);

            if (xmlStyles != null)
            {
                foreach (XmlNode style in xmlStyles)
                {
                    if (style.Attributes != null)
                        styles[style.Attributes["styleId", WordNamespace].Value] = style;
                }
            }

            var headings = new Stack<KeyValuePair<int, TreeNode>>();

            foreach (var block in EnumerateBlocks(body))
            {
                var text = ReadText(block, namespaces);
                var label = Regex.Replace(text, @"\s+", " ").Trim();
                
                if (label.Length == 0)
                {
                    if (block.SelectSingleNode(".//w:drawing | .//w:pict | .//w:object", namespaces) != null)
                        label = "[Image / object]";
                    else if (block.LocalName != "tbl")
                        continue; // Empty layout paragraphs do not need a selectable node.
                }

                if (block.LocalName == "tbl")
                    label = "[Table] " + label;
                
                var level = GetHeadingLevel(block, styles, namespaces);
                
                if (level >= 0)
                {
                    while (headings.Count > 0 && headings.Peek().Key >= level)
                        headings.Pop();
                }

                // Keep the original element, including formatting and image references, for export.
                var node = new TreeNode(label.Length > 160 ? label.Substring(0, 157) + "..." : label)
                {
                    Tag = block
                };
                
                var parentNodes = headings.Count == 0 ? treeView1.Nodes : headings.Peek().Value.Nodes;
                
                parentNodes.Add(node);
                
                if (level >= 0)
                    headings.Push(new KeyValuePair<int, TreeNode>(level, node));
            }
        }
        finally
        {
            treeView1.EndUpdate();
        }
    }

    internal static IEnumerable<XmlNode> EnumerateBlocks(XmlNode container)
    {
        foreach (XmlNode child in container.ChildNodes)
        {
            if (child.NamespaceURI != WordNamespace)
                continue;
            
            if (child.LocalName is "p" or "tbl")
            {
                // A table is one selectable unit; its cell paragraphs are not added again.
                yield return child;
            }
            else if (child.LocalName is "sdt" or "sdtContent" or "customXml" or "ins")
            {
                foreach (var block in EnumerateBlocks(child))
                    yield return block;
            }
        }
    }

    private static int GetHeadingLevel(XmlNode paragraph, Dictionary<string, XmlNode> styles, XmlNamespaceManager namespaces)
    {
        var properties = paragraph.SelectSingleNode("w:pPr", namespaces);
        var outline = properties?.SelectSingleNode("w:outlineLvl/@w:val", namespaces);
        var styleId = properties?.SelectSingleNode("w:pStyle/@w:val", namespaces)?.Value;
        var visited = new HashSet<string>(StringComparer.Ordinal);

        // EASA uses custom styles such as Heading4IR, some inheriting their outline level.
        while (outline == null && styleId != null && visited.Add(styleId) && styles.TryGetValue(styleId, out var style))
        {
            outline = style.SelectSingleNode("w:pPr/w:outlineLvl/@w:val", namespaces);
            styleId = style.SelectSingleNode("w:basedOn/@w:val", namespaces)?.Value;
        }
        
        return outline != null && int.TryParse(outline.Value, out var level) && level >= 0 && level < 9
            ? level : -1;
    }

    private static string ReadText(XmlNode block, XmlNamespaceManager namespaces)
    {
        var text = new StringBuilder();

        foreach (XmlNode part in block.SelectNodes(".//w:t | .//w:tab | .//w:br | .//w:cr | .//w:p", namespaces)!)
        {
            switch (part.LocalName)
            {
                case "t": text.Append(part.InnerText);
                    break;
                case "tab": text.Append('\t');
                    break;
                default:
                    if (text.Length > 0)
                        text.AppendLine();

                    break;
            }
        }

        return text.ToString().Trim();
    }

    private void TreeViewAfterCheck(object sender, TreeViewEventArgs e)
    {
        if (_updatingChecks)
            return;

        _updatingChecks = true;
        treeView1.BeginUpdate();

        try
        {
            // Checking a chapter selects all its content; individual children can then be excluded.
            SetDescendantChecks(e.Node, e.Node.Checked);
        }
        finally
        {
            _updatingChecks = false;
            treeView1.EndUpdate();
        }
    }

    private static void SetDescendantChecks(TreeNode parent, bool isChecked)
    {
        foreach (TreeNode child in parent.Nodes)
        {
            child.Checked = isChecked;
            SetDescendantChecks(child, isChecked);
        }
    }

    private void TreeViewAfterSelect(object sender, TreeViewEventArgs e)
    {
        var namespaces = new XmlNamespaceManager(Document.NameTable);
        namespaces.AddNamespace("w", WordNamespace);
        var text = new StringBuilder();
        AppendNodeText(e.Node, text, namespaces);
        textBox1.Text = text.ToString().TrimEnd();
    }

    private static void AppendNodeText(TreeNode node, StringBuilder text, XmlNamespaceManager namespaces)
    {
        if (node.Tag is XmlNode block)
            text.AppendLine(ReadText(block, namespaces));

        foreach (TreeNode child in node.Nodes)
            AppendNodeText(child, text, namespaces);
    }

    private void exportSelectedToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var selectedBlocks = new HashSet<XmlNode>();
        CollectCheckedBlocks(treeView1.Nodes, selectedBlocks);

        if (selectedBlocks.Count == 0)
        {
            MessageBox.Show(this, @"Select at least one paragraph or chapter to export.", @"Export XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var x = new SaveFileDialog();
        x.Title = @"Save XML File";
        x.Filter = @"XML files (*.xml)|*.xml|All files (*.*)|*.*";
        x.DefaultExt = "xml";
        x.AddExtension = true;
        x.FileName = "Selected rules.xml";

        if (x.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            UseWaitCursor = true;
            var export = WordXmlExporter.Create(Document, selectedBlocks);
            WordXmlExporter.Save(export, x.FileName);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or XmlException or InvalidOperationException)
        {
            MessageBox.Show(this, @"The XML file could not be exported.\n\n" + exception.Message, @"Export XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private static void CollectCheckedBlocks(TreeNodeCollection nodes, ISet<XmlNode> selectedBlocks)
    {
        foreach (TreeNode node in nodes)
        {
            // Check every node independently: a checked heading may have unchecked children.
            if (node.Checked && node.Tag is XmlNode block)
                selectedBlocks.Add(block);

            CollectCheckedBlocks(node.Nodes, selectedBlocks);
        }
    }

    private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
    {
        MessageBox.Show(this,
            @"This program allows you to select chapters from the Easy Access Rules for Air Operations from the European Authority for aviation safety.

Written by Anders Hesselbom.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
