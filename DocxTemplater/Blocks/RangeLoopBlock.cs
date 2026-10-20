using System.Globalization;
using DocumentFormat.OpenXml;
using DocxTemplater.Schema;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Blocks
{
    internal class RangeLoopBlock : ContentBlock
    {
        private readonly string m_indexName;
        private readonly string m_countVariable;

        public RangeLoopBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
            (m_indexName, m_countVariable) = RangeLoopCount.Parse(startMatch.Variable);
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            int count = 0;
            try
            {
                count = RangeLoopCount.Resolve(models, m_countVariable, m_context.ProcessSettings);
            }
            catch (OpenXmlTemplateException e) when (m_context.ProcessSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
            {
                if (m_context.ProcessSettings.BindingErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_context.VariableReplacer.AddError(e.Message);
                }
            }

            // Iterate backwards because InsertAfterSelf inserts elements immediately after the insertion point,
            // effectively reversing the output order if not iterated backward.
            for (int j = count - 1; j >= 0; j--)
            {
                using var loopScope = models.OpenScope();
                loopScope.AddVariable(m_indexName, j);
                base.Expand(models, parentNode);
            }
        }

        public override string ToString()
        {
            return $"RangeLoopBlock: {m_indexName}:{m_countVariable}";
        }

        public override void CollectSchema(SchemaBuilder builder)
        {
            // The count side can be either an int literal or a path on the model; only declare paths.
            if (!int.TryParse(m_countVariable, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
            {
                SchemaExpressionParser.Extract(m_countVariable, builder.DeclareScalar);
            }
            // Range loops introduce a scope level for dot-resolution but no named item.
            builder.OpenTransparentScope();
            try
            {
                base.CollectSchema(builder);
            }
            finally
            {
                builder.CloseScope();
            }
        }
    }
}
