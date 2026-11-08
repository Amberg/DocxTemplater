using System;
using System.Collections.Generic;

namespace DocxTemplater.Text
{
    /// <summary>
    /// What a render needs of a text template beyond the models: the text, its syntax tree and the ranges removed by
    /// <see cref="TextTemplateSettings.IgnoreLineBreaksAroundTags"/>. Immutable, so concurrent renders can share it.
    /// </summary>
    internal sealed class ParsedTextTemplate
    {
        public ParsedTextTemplate(string text, TemplateSyntaxTree tree, bool ignoreLineBreaksAroundTags)
        {
            Text = text;
            Tree = tree;
            Helper = new SyntaxTreeHelper(tree);
            SkippedRanges = ignoreLineBreaksAroundTags
                ? SyntaxTreeHelper.LineBreakRangesToSkip(text, tree.Matches)
                : Array.Empty<(int, int)>();
        }

        public string Text { get; }

        public TemplateSyntaxTree Tree { get; }

        public SyntaxTreeHelper Helper { get; }

        public IReadOnlyList<(int Start, int End)> SkippedRanges { get; }
    }
}
