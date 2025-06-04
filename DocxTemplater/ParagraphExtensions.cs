using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater
{
    internal static class ParagraphExtensions
    {
        // Use a static dictionary to track paragraphs containing only blocks
        private static readonly Dictionary<Paragraph, bool> ParagraphBlockMarkers = new Dictionary<Paragraph, bool>();

        /// <summary>
        /// Marks a paragraph as containing only template blocks (like loop beginnings/endings, conditionals)
        /// </summary>
        public static void MarkAsContainingOnlyBlocks(this Paragraph paragraph)
        {
            ParagraphBlockMarkers[paragraph] = true;
        }

        /// <summary>
        /// Checks if a paragraph is marked as containing only template blocks
        /// </summary>
        public static bool IsMarkedAsContainingOnlyBlocks(this Paragraph paragraph)
        {
            return ParagraphBlockMarkers.ContainsKey(paragraph);
        }

        /// <summary>
        /// Removes the marker indicating a paragraph contains only template blocks
        /// </summary>
        public static void RemoveContainsOnlyBlocksMarker(this Paragraph paragraph)
        {
            ParagraphBlockMarkers.Remove(paragraph);
        }

        /// <summary>
        /// Cleans up all paragraph markers from the dictionary
        /// </summary>
        public static void CleanupAllMarkers()
        {
            ParagraphBlockMarkers.Clear();
        }
    }
} 