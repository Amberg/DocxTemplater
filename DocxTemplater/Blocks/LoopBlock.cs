using System;
using System.Collections;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Schema;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Blocks
{
    internal class LoopBlock : ContentBlock
    {
        public static string LoopIndexVariable = "_Idx";
        public static string LoopLengthVariable = "_Length";

        private readonly string m_collectionName;
        private readonly ListNumberingRestarter m_listNumberingRestarter;

        public LoopBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
            m_collectionName = startMatch.Variable;
            m_listNumberingRestarter = new ListNumberingRestarter(this, context);
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            object model = null;
            try
            {
                model = models.GetValue(m_collectionName);
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

            if (model is IEnumerable enumerable)
            {
                var items = enumerable.Cast<object>().Reverse().ToList();
                int counter = items.Count;
                foreach (var item in items)
                {
                    using var loopScope = models.OpenScope();
                    loopScope.AddVariable(m_collectionName, item);
                    loopScope.AddVariable($"{m_collectionName}.{LoopIndexVariable}", counter);
                    loopScope.AddVariable($"{m_collectionName}.{LoopLengthVariable}", items.Count);
                    var iteration = m_listNumberingRestarter.BeginIteration(parentNode, isFirstInDocument: counter == 1);
                    base.Expand(models, parentNode);
                    m_listNumberingRestarter.EndIteration(iteration);
                    counter--;
                }
            }
            else if (model != null)
            {
                if (m_context.VariableReplacer.ProcessSettings.BindingErrorHandling == BindingErrorHandling.ThrowException)
                {
                    throw OpenXmlTemplateException.Create(m_context.ProcessSettings, TemplateErrorCode.NotEnumerable, m_collectionName, model.GetType().FullName);
                }
                else
                {
                    m_context.VariableReplacer.AddError(TemplateErrorCode.NotEnumerable, m_collectionName, model.GetType().FullName);
                }
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }

        public override void CollectSchema(SchemaBuilder builder)
        {
            // m_collectionName can be dot-prefixed (e.g. `.subitems` for a nested loop on the outer item).
            // The SchemaBuilder needs the absolute key against the current scope stack to register the item scope.
            var absoluteKey = builder.ResolveAbsoluteScopeKey(m_collectionName);
            var itemSchema = builder.DeclareCollection(m_collectionName);
            if (itemSchema != null)
            {
                builder.OpenItemScope(absoluteKey, itemSchema);
            }
            else
            {
                builder.OpenTransparentScope();
            }
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