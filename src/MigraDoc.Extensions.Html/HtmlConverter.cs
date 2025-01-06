using HtmlAgilityPack;
using MigraDoc.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace MigraDoc.Extensions.Html
{
    public class HtmlConverter : IConverter
    {
        private IDictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>> nodeHandlers
            = new Dictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>>();
           
        public HtmlConverter()
        {
            AddDefaultNodeHandlers();
        }

        public IDictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>> NodeHandlers
        {
            get
            {
                return nodeHandlers;
            }
        }

        public Action<DocumentObject> Convert(string contents)
        {
            return section => ConvertHtml(contents, section);
        }

        private void ConvertHtml(string html, DocumentObject section)
        {
            if (string.IsNullOrEmpty(html))
            {
                throw new ArgumentNullException("html");
            }

            if (section == null)
            {
                throw new ArgumentNullException("section");
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            ConvertHtmlNodes(doc.DocumentNode.ChildNodes, section);
        }

        private void ConvertHtmlNodes(HtmlNodeCollection nodes, DocumentObject section, DocumentObject current = null)
        {
            DocumentObject previous = null;
            foreach (var node in nodes)
            {
                Func<HtmlNode, DocumentObject, DocumentObject> nodeHandler;
                if (nodeHandlers.TryGetValue(node.Name, out nodeHandler))
                {
                    if (node.PreviousSibling != null && node.PreviousSibling.Name == "#text" && previous != null && previous is Paragraph)
                    {
                        current = (previous as Paragraph);
                    }
                    // pass the current container or section
                    var result = nodeHandler(node, current ?? section);
                    
                    if (node.HasChildNodes)
                    {
                        ConvertHtmlNodes(node.ChildNodes, section, result);
                    }
                    if (node.NextSibling != null && node.NextSibling.Name == "span")
                    {
                        // MISSING: mit vorherigem Paragraph zusammenfügen / Text anhängen
           //             ConvertHtmlNodes(new HtmlNodeCollection() { node.NextSibling }, section, result)};
                    }
                    previous = result;
                }
                else
                {
                    if (node.HasChildNodes)
                    {
                        ConvertHtmlNodes(node.ChildNodes, section, current);
                    }
                }
            }
        }
        
        private void AddDefaultNodeHandlers()
        {
            // Block Elements
            
            // could do with a predicate/regex matcher so we could just use one handler for all headings
            nodeHandlers.Add("h1", AddHeading);
            nodeHandlers.Add("h2", AddHeading);
            nodeHandlers.Add("h3", AddHeading);
            nodeHandlers.Add("h4", AddHeading);
            nodeHandlers.Add("h5", AddHeading);
            nodeHandlers.Add("h6", AddHeading);

            nodeHandlers.Add("p", AddParagraph);
            nodeHandlers.Add("div", AddParagraph);

            // Inline Elements

            nodeHandlers.Add("strong", (node, parent) => AddFormattedText(node, parent, TextFormat.Bold));
            nodeHandlers.Add("bold", (node, parent) => AddFormattedText(node, parent, TextFormat.Bold));
            nodeHandlers.Add("i", (node, parent) => AddFormattedText(node, parent, TextFormat.Italic));
            nodeHandlers.Add("em", (node, parent) => AddFormattedText(node, parent, TextFormat.Italic));
            nodeHandlers.Add("u", (node, parent) => AddFormattedText(node, parent, TextFormat.Underline));
            nodeHandlers.Add("a", (node, parent) =>
            {
                return GetParagraph(parent).AddHyperlink(node.GetAttributeValue("href", ""), HyperlinkType.Web);
            });
            nodeHandlers.Add("span", (node, parent) =>
            {
                if (node.InnerText == "&nbsp;") return parent;
                var p = GetParagraph(parent).AddFormattedText(TextFormat.NoUnderline);
                var cssClass = node.Attributes["class"];
                if (cssClass != null && !string.IsNullOrEmpty(cssClass.Value) && cssClass.Value != "style_color_0_0_0")
                {
                    p.Style = cssClass.Value;
                }
                return p;
            });
            nodeHandlers.Add("hr", (node, parent) => GetParagraph(parent).SetStyle("HorizontalRule"));
            nodeHandlers.Add("br", (node, parent) => {
                if (node.NextSibling == null) return parent; // do not add a line break if last element in parent group
                if (node.NextSibling.Name == "ul" || node.NextSibling.Name == "ol") return parent; // do not add a line break if list is next element
                if (parent is FormattedText)
                {
                    // inline elements can contain line breaks
                    ((FormattedText)parent).AddLineBreak();
                    return parent;
                }
                if (parent is Paragraph)
                {
                    // inline elements can contain line breaks
                    ((Paragraph)parent).AddLineBreak();
                    return parent;
                }
                if (parent is Section)
                {
                    // section add empty line break => not good/results in error?
                    if(node.NextSibling.Name != "br") return parent;
                }       
   
                var paragraph = GetParagraph(parent);
                paragraph.AddLineBreak();
                return paragraph;
            });

            nodeHandlers.Add("li", (node, parent) =>
            {
                var listStyle = node.ParentNode.Name.ToLower().Trim() == "ul"
                    ? "UnorderedList"
                    : "OrderedList";

                var isFirst = node.ParentNode.Elements("li").First() == node;
                var isLast = node.ParentNode.Elements("li").Last() == node;
                
                Section section;
                if (parent is Paragraph)
                {
                    section = parent.Section;
                }
                else // if (parent is Section)
                {
                    section = (Section) parent;
                }
                // if this is the first item add the ListStart paragraph
                if (isFirst)
                {
                    section.AddParagraph().SetStyle("ListStart");
                }

                Paragraph listItem = section.AddParagraph().SetStyle(listStyle);

                // disable continuation if this is the first list item
                listItem.Format.ListInfo.ContinuePreviousList = !isFirst;
                // add list style/layout                
                if (listStyle == "OrderedList")
                {
                    listItem.Format.ListInfo.ListType = ListType.NumberList1;
                }
                else
                {
                    listItem.Format.ListInfo.ListType = ListType.BulletList1;
                }

                // if the this is the last item add the ListEnd paragraph
                if (isLast)
                {
                    section.AddParagraph().SetStyle("ListEnd");
                }

                return listItem;
            });

            nodeHandlers.Add("#text", (node, parent) =>
            {
                // remove line breaks
                var innerText = node.InnerText.Replace("\r", "").Replace("\n", "");

                if (string.IsNullOrWhiteSpace(innerText))
                {
                    return parent;
                }

                // decode escaped HTML
                innerText = WebUtility.HtmlDecode(innerText);
                
                // text elements must be wrapped in a paragraph but this could also be FormattedText or a Hyperlink!!
                // this needs some work
                if (parent is FormattedText)
                {
                    return ((FormattedText)parent).AddText(innerText);
                }
                if (parent is Hyperlink)
                {
                    return ((Hyperlink)parent).AddText(innerText);
                }

                // otherwise a section or paragraph
                var res = GetParagraph(parent);
                res.AddText(innerText);
                return res;
            });
        }

        private DocumentObject AddParagraph(HtmlNode node, DocumentObject parent)
        {
            Paragraph res;
            res = (parent is Paragraph) ? (parent.Section.AddParagraph()) : ((Section) parent).AddParagraph();
            var styleAttrib = node.Attributes["style"];
            if (styleAttrib != null && !string.IsNullOrEmpty(styleAttrib.Value))
            {
                string styleAttibute = styleAttrib.Value.ToLower().Trim();
                const string textAlignMatch = "text-align:";
                if (styleAttibute.Contains(textAlignMatch)) // match paragraph align
                {
                    int textAlignFrom = styleAttibute.IndexOf(textAlignMatch, StringComparison.InvariantCultureIgnoreCase) + textAlignMatch.Length;
                    int textAlignTo = styleAttibute.IndexOf(";", textAlignFrom, StringComparison.InvariantCultureIgnoreCase);
                    if (textAlignTo < textAlignFrom) textAlignTo = styleAttibute.Length;
                    string textAlign = styleAttrib.Value.Substring(textAlignFrom, textAlignTo - textAlignFrom).Trim();
                    SetParagraphAlignment(res, textAlign);
                }
            }
            return res;
        }

        private static void SetParagraphAlignment(Paragraph res, string textAlign)
        {
            switch (textAlign.ToLower())
            {
                case "center":
                    res.Format.Alignment = ParagraphAlignment.Center;
                    break;
                case "left":
                    res.Format.Alignment = ParagraphAlignment.Left;
                    break;
                case "right":
                    res.Format.Alignment = ParagraphAlignment.Right;
                    break;
                case "justify":
                    res.Format.Alignment = ParagraphAlignment.Justify;
                    break;
            }
        }

        private static DocumentObject AddFormattedText(HtmlNode node, DocumentObject parent, TextFormat format)
        {
            var formattedText = parent as FormattedText;
            if (formattedText != null)
            {
                return formattedText.Format(format);
            }

            // otherwise parent is paragraph or section
            return GetParagraph(parent).AddFormattedText(format);
        }

        private static DocumentObject AddHeading(HtmlNode node, DocumentObject parent)
        {
            return ((Section)parent).AddParagraph().SetStyle("Heading" + node.Name[1]);
        }

        private static Paragraph GetParagraph(DocumentObject parent)
        {
            return parent as Paragraph ?? ((Section)parent).AddParagraph();
        }

        private static Paragraph AddParagraphWithStyle(DocumentObject parent, string style) 
        {
            return ((Section)parent).AddParagraph().SetStyle(style);
        }
    }
}
