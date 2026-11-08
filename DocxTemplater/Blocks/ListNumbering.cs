using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Blocks
{
    internal static class ListNumbering
    {
        /// <summary>
        /// Restarts the lists marked as "Restart numbering" in Word (a w:num with a w:startOverride) in one loop iteration -
        /// the elements after <paramref name="anchor"/> up to <paramref name="end"/> - by giving them a copy of their w:num
        /// (https://github.com/Amberg/DocxTemplater/issues/162). The list the iteration starts with is kept:
        /// then the loop writes the items of that list.
        /// </summary>
        public static void RestartLists(ITemplateProcessingContext context, OpenXmlElement anchor, OpenXmlElement end)
        {
            var numbering = context.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            var paragraphs = new List<Paragraph>();
            for (var e = anchor?.NextSibling(); numbering != null && e != null && e != end; e = e.NextSibling())
            {
                paragraphs.AddRange(e is Paragraph p ? [p] : e.Descendants<Paragraph>());
            }
            var keep = NumId(paragraphs.FirstOrDefault(p => NumId(p) != null || !string.IsNullOrWhiteSpace(p.InnerText)))?.Val?.Value;
            var copies = new Dictionary<int, int>();
            foreach (var numId in paragraphs.Select(NumId).Where(x => x?.Val != null && x.Val.Value != keep))
            {
                if (!copies.TryGetValue(numId.Val, out var newId))
                {
                    var num = numbering.Elements<NumberingInstance>().FirstOrDefault(x => x.NumberID == numId.Val && x.Descendants<StartOverrideNumberingValue>().Any());
                    newId = numId.Val;
                    if (num != null)
                    {
                        var copy = (NumberingInstance)num.CloneNode(true);
                        copy.NumberID = newId = numbering.Elements<NumberingInstance>().Max(x => x.NumberID?.Value ?? 0) + 1;
                        copy.RemoveAttribute("durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                        numbering.InsertAfterLastChildOfSameType(copy);
                    }
                    copies[numId.Val] = newId;
                }
                numId.Val = newId;
            }
        }

        private static NumberingId NumId(Paragraph p)
        {
            return p?.ParagraphProperties?.NumberingProperties?.NumberingId;
        }
    }
}
