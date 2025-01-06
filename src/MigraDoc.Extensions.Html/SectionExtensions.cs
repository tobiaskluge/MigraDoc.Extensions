using MigraDoc.DocumentObjectModel;
using MigraDoc.Extensions.Html;
using System;

namespace MigraDoc.Extensions.Html
{
    public static class SectionExtensions
    {
        public static Section AddHtml(this Section section, string html)
        {
            if (string.IsNullOrEmpty(html)) return null;
            return section.Add(html, new HtmlConverter());
        }
    }
}
