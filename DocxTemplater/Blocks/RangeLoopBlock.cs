using System;
using System.Collections;
using System.Linq;
using DocumentFormat.OpenXml;
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
                            throw new OpenXmlTemplateException(
                                $"'{m_countVariable}' is not an integer - its value '{stringValue}' cannot be parsed to an integer");
                        }
                        else
                        {
                            m_context.VariableReplacer.AddError(
                                $"'{m_countVariable}' is not an integer - its value '{stringValue}' cannot be parsed to an integer");
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
                            if (!Enumerable.TryGetNonEnumeratedCount(enumerableObjects, out count))
                            {
                                count = enumerableObjects.Count();
                            }
                        }
                        else
                        {
                            if (m_context.ProcessSettings.BindingErrorHandling == BindingErrorHandling.ThrowException)
                            {
                                throw new OpenXmlTemplateException(
                                    $"'{m_countVariable}' is not an integer or enumerable - it is of type {model.GetType().FullName}");
                            }
                            else
                            {
                                m_context.VariableReplacer.AddError(
                                    $"'{m_countVariable}' is not an integer or enumerable - it is of type {model.GetType().FullName}");
                            }
                        }
                    }
                }
            }

            if (count < 0)
            {
                count = 0;
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
    }
}
