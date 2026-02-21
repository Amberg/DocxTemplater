using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Blocks
{
    internal class CaseBlock : ContentBlock
    {
        public string MatchExpression { get; }
        public bool IsDefault { get; }
        public bool IsMatched { get; set; }

        public CaseBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
            var matchArg = startMatch.Variable.Trim();
            if (matchArg.StartsWith("case:", StringComparison.OrdinalIgnoreCase) || matchArg.StartsWith("c:", StringComparison.OrdinalIgnoreCase))
            {
                MatchExpression = matchArg[(matchArg.IndexOf(':') + 1)..].Trim();
                IsDefault = false;
            }
            else if (matchArg.Equals("default", StringComparison.OrdinalIgnoreCase) || matchArg.Equals("d", StringComparison.OrdinalIgnoreCase))
            {
                MatchExpression = null;
                IsDefault = true;
            }
            else
            {
                throw new OpenXmlTemplateException($"Invalid case block syntax: {startMatch.Variable}");
            }
        }

        public override string ToString()
        {
            return $"CaseBlock: {(IsDefault ? "default" : MatchExpression)}";
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            if (!IsNestedUnderSwitch())
            {
                var caseDescription = IsDefault ? "default" : MatchExpression;
                throw new OpenXmlTemplateException(
                    $"The '{{#case}}'/'{{#default}}' block ('{caseDescription}') must be nested inside a '{{#switch}}' block.");
            }

            if (IsMatched)
            {
                base.Expand(models, parentNode);
            }
        }

        private bool IsNestedUnderSwitch()
        {
            var current = ParentBlock;
            while (current != null)
            {
                if (current is SwitchBlock)
                {
                    return true;
                }
                current = current.ParentBlock;
            }
            return false;
        }
    }
}
