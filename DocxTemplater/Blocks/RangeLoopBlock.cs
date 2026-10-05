using System;
using System.Collections;
using System.Globalization;
using System.Linq;
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
            var varName = startMatch.Variable ?? string.Empty;
            if (varName.Contains(':'))
            {
                var parts = varName.Split(':');
                m_indexName = parts[0].Trim();
                m_countVariable = parts[1].Trim();
            }
            else
            {
                m_indexName = "Index";
                m_countVariable = varName.Trim();
            }
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            // The count can be an integer literal ({{@i:3}}) - resolve it directly instead of looking it up on the model.
            if (int.TryParse(m_countVariable, NumberStyles.Integer, CultureInfo.InvariantCulture, out int literalCount))
            {
                ExpandIterations(models, parentNode, literalCount);
                return;
            }

            object model = null;
            try
            {
                model = models.GetValue(m_countVariable);
            }
            catch (OpenXmlTemplateException e) when (m_context.ProcessSettings.BindingErrorHandling !=
                                                     BindingErrorHandling.ThrowException)
            {
                if (m_context.ProcessSettings.BindingErrorHandling ==
                    BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_context.VariableReplacer.AddError(e.Message);
                }
            }

            int count = 0;
            if (model != null)
            {
                if (model is string stringValue)
                {
                    if (!int.TryParse(stringValue, out count))
                    {
                        if (m_context.ProcessSettings.BindingErrorHandling == BindingErrorHandling.ThrowException)
                        {
                            throw OpenXmlTemplateException.Create(m_context.ProcessSettings, TemplateErrorCode.RangeCountNotNumeric, m_countVariable, stringValue);
                        }
                        else
                        {
                            m_context.VariableReplacer.AddError(TemplateErrorCode.RangeCountNotNumeric, m_countVariable, stringValue);
                        }
                    }
                }
                else
                {
                    try
                    {
                        count = Convert.ToInt32(model);
                    }
                    catch (Exception)
                    {
                        if (model is IEnumerable enumerable)
                        {
                            var enumerableObjects = enumerable.Cast<object>();
                            if (!enumerableObjects.TryGetNonEnumeratedCount(out count))
                            {
                                count = enumerableObjects.Count();
                            }
                        }
                        else
                        {
                            if (m_context.ProcessSettings.BindingErrorHandling == BindingErrorHandling.ThrowException)
                            {
                                throw OpenXmlTemplateException.Create(m_context.ProcessSettings, TemplateErrorCode.RangeCountNotNumericOrEnumerable, m_countVariable, model.GetType().FullName);
                            }
                            else
                            {
                                m_context.VariableReplacer.AddError(TemplateErrorCode.RangeCountNotNumericOrEnumerable, m_countVariable, model.GetType().FullName);
                            }
                        }
                    }
                }
            }

            ExpandIterations(models, parentNode, count);
        }

        private void ExpandIterations(IModelLookup models, OpenXmlElement parentNode, int count)
        {
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
