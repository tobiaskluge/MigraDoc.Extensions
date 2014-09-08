using System;
using MigraDoc.DocumentObjectModel;

namespace MigraDoc.Extensions.Html
{
    public static class ParagraphExtensions
    {
        public static Paragraph AddHtml(this Paragraph paragraph, string html)
        {
            return paragraph.Add(html, new HtmlConverter());
        }
    }
}
