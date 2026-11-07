using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Blocks
{
    /// <summary>
    /// Restarts numbered lists in every iteration of a loop (https://github.com/Amberg/DocxTemplater/issues/162).
    /// <para>
    /// Word marks a list as "Restart numbering" with a dedicated <c>w:num</c> that carries a
    /// <c>w:lvlOverride/w:startOverride</c>. A loop clones the paragraphs, but all clones still reference the
    /// same <c>w:num</c> - for Word this is one single list that continues over all iterations.
    /// To restart the list, every iteration (except the first in document order) gets its own copy of that
    /// <c>w:num</c>.
    /// </para>
    /// <para>
    /// Only a loop whose body does not start with an item of that list restarts it: in
    /// <c>{{#Items}}1. {{.Name}}{{/Items}}</c> the iterations produce the items of one list and must keep it,
    /// while in <c>{{#Groups}}Title, list{{/Groups}}</c> every group starts a new list.
    /// </para>
    /// </summary>
    internal sealed class ListNumberingRestarter
    {
        private readonly ContentBlock m_block;
        private readonly ITemplateProcessingContext m_context;
        private HashSet<int> m_numIdsToRestart;

        public ListNumberingRestarter(ContentBlock block, ITemplateProcessingContext context)
        {
            m_block = block;
            m_context = context;
        }

        /// <summary>
        /// Call before an iteration is expanded. Returns the element after which the iteration is inserted and the
        /// element following it, or null when the items of this iteration keep their list. The iteration that ends
        /// up first in the document keeps the list of the template.
        /// </summary>
        public (OpenXmlElement insertionElement, OpenXmlElement next)? BeginIteration(OpenXmlElement parentNode, bool isFirstInDocument)
        {
            if (isFirstInDocument || GetNumIdsToRestart().Count == 0)
            {
                return null;
            }
            var insertionElement = m_block.GetInsertionPointElement(parentNode);
            if (insertionElement == null)
            {
                return null;
            }
            return (insertionElement, insertionElement.NextSibling());
        }

        /// <summary>
        /// Moves the list items of the iteration inserted since <see cref="BeginIteration"/> to a new restarted list.
        /// </summary>
        public void EndIteration((OpenXmlElement insertionElement, OpenXmlElement next)? iteration)
        {
            if (iteration == null)
            {
                return;
            }
            var numbering = m_context.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null)
            {
                return;
            }
            var (insertionElement, next) = iteration.Value;
            var newNumIds = new Dictionary<int, int>();
            for (var element = insertionElement.NextSibling(); element != null && element != next; element = element.NextSibling())
            {
                foreach (var numberingId in element.Descendants<NumberingId>())
                {
                    if (numberingId.Val == null || !m_numIdsToRestart.Contains(numberingId.Val.Value))
                    {
                        continue;
                    }
                    if (!newNumIds.TryGetValue(numberingId.Val.Value, out var newNumId))
                    {
                        newNumId = CreateRestartedCopy(numbering, numberingId.Val.Value);
                        newNumIds[numberingId.Val.Value] = newNumId;
                    }
                    numberingId.Val = newNumId;
                }
            }
        }

        private static int CreateRestartedCopy(Numbering numbering, int numId)
        {
            var original = numbering.Elements<NumberingInstance>().First(x => x.NumberID?.Value == numId);
            var copy = (NumberingInstance)original.CloneNode(true);
            copy.NumberID = numbering.Elements<NumberingInstance>().Max(x => x.NumberID?.Value ?? 0) + 1;
            copy.RemoveAttribute("durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            numbering.InsertAfterLastChildOfSameType(copy);
            return copy.NumberID.Value;
        }

        private HashSet<int> GetNumIdsToRestart()
        {
            if (m_numIdsToRestart != null)
            {
                return m_numIdsToRestart;
            }
            m_numIdsToRestart = new HashSet<int>();
            var numbering = m_context.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null)
            {
                return m_numIdsToRestart;
            }

            // lists marked as "Restart numbering" in the template
            var restartingNumIds = numbering.Elements<NumberingInstance>()
                .Where(x => x.NumberID != null && x.Elements<LevelOverride>().Any(o => o.StartOverrideNumberingValue != null))
                .Select(x => x.NumberID.Value)
                .ToHashSet();
            if (restartingNumIds.Count == 0)
            {
                return m_numIdsToRestart;
            }

            var paragraphs = TemplateParagraphs(m_block).ToList();
            var firstParagraph = paragraphs.FirstOrDefault(x => GetNumId(x) != null || !string.IsNullOrWhiteSpace(x.InnerText));
            var numIdOfFirstParagraph = firstParagraph == null ? null : GetNumId(firstParagraph);
            foreach (var numId in paragraphs.Select(GetNumId))
            {
                if (numId != null && numId != numIdOfFirstParagraph && restartingNumIds.Contains(numId.Value))
                {
                    m_numIdsToRestart.Add(numId.Value);
                }
            }
            return m_numIdsToRestart;
        }

        private static int? GetNumId(Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        }

        /// <summary>
        /// The paragraphs of the body of the block in document order - including the content of the child blocks at
        /// the position of their insertion point.
        /// </summary>
        private static IEnumerable<Paragraph> TemplateParagraphs(ContentBlock block)
        {
            return block.Content.SelectMany(x => TemplateParagraphs(block, x));
        }

        private static IEnumerable<Paragraph> TemplateParagraphs(ContentBlock block, OpenXmlElement element)
        {
            var child = block.ChildBlocks.FirstOrDefault(c => c.IsInsertionPoint(element));
            if (child != null)
            {
                foreach (var paragraph in TemplateParagraphs(child))
                {
                    yield return paragraph;
                }
            }
            if (element is Paragraph p)
            {
                yield return p;
                yield break;
            }
            foreach (var childElement in element.ChildElements)
            {
                foreach (var paragraph in TemplateParagraphs(block, childElement))
                {
                    yield return paragraph;
                }
            }
        }
    }
}
