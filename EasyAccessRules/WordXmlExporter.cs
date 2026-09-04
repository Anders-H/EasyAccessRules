#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace EasyAccessRules;

public static class WordXmlExporter
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static XmlDocument Create(XmlDocument source, ISet<XmlNode> selectedBlocks)
    {
        if (selectedBlocks.Count == 0)
            throw new InvalidOperationException("Select at least one paragraph or chapter to export.");

        var namespaces = new XmlNamespaceManager(source.NameTable);
        namespaces.AddNamespace("w", WordNamespace);
        namespaces.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");
        const string bodyPath = "/pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body | /w:document/w:body";
        var sourceBody = source.SelectSingleNode(bodyPath, namespaces) ?? throw new InvalidOperationException("No Word document body was found in the XML file.");

        // Preserve all package parts, including binary data, relationships, styles and footnotes.
        // Work on a copy so another selection can be exported from the same open document.
        var result = (XmlDocument)source.CloneNode(true);
        var body = result.SelectSingleNode(bodyPath, namespaces)!;

        while (body.FirstChild != null)
            body.RemoveChild(body.FirstChild);

        var inheritedReferences = new Dictionary<string, XmlElement>(StringComparer.Ordinal);
        var sectionHasContent = false;
        XmlElement? lastSectionBreak = null;
        var copiedCount = 0;

        foreach (var block in MainWindow.EnumerateBlocks(sourceBody))
        {
            if (selectedBlocks.Contains(block))
            {
                var copy = (XmlElement)result.ImportNode(block, true);
                var inlineSection = copy.SelectSingleNode("w:pPr/w:sectPr", namespaces);
                inlineSection?.ParentNode?.RemoveChild(inlineSection);

                // Export paragraphs directly, without content-control bindings or placeholders
                // belonging to deselected paragraphs. Each selected table remains intact.
                body.AppendChild(copy);
                sectionHasContent = true;
                copiedCount++;
            }

            var section = block.SelectSingleNode("w:pPr/w:sectPr", namespaces) as XmlElement;

            if (section != null)
            {
                FinishSection(section);
                sectionHasContent = false;
            }
        }

        FinishSection(sourceBody.SelectSingleNode("w:sectPr", namespaces) as XmlElement);

        if (copiedCount != selectedBlocks.Count)
            throw new InvalidOperationException("The selection does not belong to the current document. Reopen the XML file and select the content again.");

        // Word stores the final section's properties directly at the end of the body.
        if (lastSectionBreak != null)
        {
            var section = lastSectionBreak.SelectSingleNode("w:pPr/w:sectPr", namespaces)!;
            section.ParentNode!.RemoveChild(section);
            body.RemoveChild(lastSectionBreak);
            body.AppendChild(section);
        }

        MakeFieldsStatic(body, namespaces);
        RemoveIncompleteBookmarks(body, namespaces);
        return result;

        void FinishSection(XmlElement? sourceSection)
        {
            // Missing header/footer references inherit from earlier sections, including sections
            // the user has excluded. Materialize those references before dropping empty sections.
            if (sourceSection != null)
            {
                foreach (XmlElement reference in sourceSection.SelectNodes("w:headerReference | w:footerReference", namespaces)!)
                    inheritedReferences[reference.LocalName + ":" + reference.GetAttribute("type", WordNamespace)] = reference;
            }

            if (!sectionHasContent)
                return;

            var section = sourceSection == null
                ? result.CreateElement("w", "sectPr", WordNamespace)
                : (XmlElement)result.ImportNode(sourceSection, true);

            foreach (XmlNode reference in section.SelectNodes("w:headerReference | w:footerReference", namespaces)!)
                section.RemoveChild(reference);

            // Header references precede footer references and the remaining section properties.
            var firstProperty = section.FirstChild;

            foreach (var name in new[] { "headerReference", "footerReference" })
            {
                foreach (var reference in inheritedReferences.Values)
                {
                    if (reference.LocalName == name)
                        section.InsertBefore(result.ImportNode(reference, true), firstProperty);
                }
            }

            var paragraph = result.CreateElement("w", "p", WordNamespace);
            var properties = result.CreateElement("w", "pPr", WordNamespace);
            properties.AppendChild(section);
            paragraph.AppendChild(properties);
            body.AppendChild(paragraph);
            lastSectionBreak = paragraph;
        }
    }

    private static void MakeFieldsStatic(XmlNode body, XmlNamespaceManager namespaces)
    {
        // TOC/REF fields can span deselected paragraphs or refer to removed bookmarks.
        // Keep their displayed result as ordinary text, without dangling field instructions.
        // PAGE fields in headers and footers are outside the body and remain dynamic.
        foreach (XmlNode fieldPart in body.SelectNodes(".//w:fldChar | .//w:instrText", namespaces)!)
            fieldPart.ParentNode!.RemoveChild(fieldPart);

        foreach (XmlNode field in body.SelectNodes(".//w:fldSimple", namespaces)!)
        {
            var parent = field.ParentNode!;

            while (field.FirstChild != null)
                parent.InsertBefore(field.FirstChild, field);

            parent.RemoveChild(field);
        }
    }

    private static void RemoveIncompleteBookmarks(XmlNode body, XmlNamespaceManager namespaces)
    {
        var starts = new HashSet<string>(StringComparer.Ordinal);
        var ends = new HashSet<string>(StringComparer.Ordinal);
        var bookmarks = body.SelectNodes(".//w:bookmarkStart | .//w:bookmarkEnd", namespaces)!;

        foreach (XmlElement bookmark in bookmarks)
        {
            var ids = bookmark.LocalName == "bookmarkStart" ? starts : ends;
            ids.Add(bookmark.GetAttribute("id", WordNamespace));
        }

        foreach (XmlElement bookmark in bookmarks)
        {
            var id = bookmark.GetAttribute("id", WordNamespace);

            if (!starts.Contains(id) || !ends.Contains(id))
                bookmark.ParentNode!.RemoveChild(bookmark);
        }
    }

    internal static void Save(XmlDocument document, string fileName)
    {
        var destination = Path.GetFullPath(fileName);
        var temporaryFile = Path.Combine(Path.GetDirectoryName(destination)!, "." + Guid.NewGuid().ToString("N") + ".tmp");

        try
        {
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                NewLineHandling = NewLineHandling.None
            };

            using (var writer = XmlWriter.Create(temporaryFile, settings))
                document.Save(writer);

            // Do not truncate an existing export if serialization or writing fails.
            if (File.Exists(destination))
                File.Replace(temporaryFile, destination, null);
            else
                File.Move(temporaryFile, destination);
        }
        finally
        {
            if (File.Exists(temporaryFile))
                File.Delete(temporaryFile);
        }
    }
}
