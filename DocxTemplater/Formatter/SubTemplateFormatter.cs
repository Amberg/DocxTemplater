using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace DocxTemplater.Formatter
{
    internal class SubTemplateFormatter : IFormatter
    {

        private readonly ModelLookup m_modelLookup;
        private readonly ProcessSettings m_settings;

        public SubTemplateFormatter(
            ModelLookup modelLookup,
            ProcessSettings settings)
        {
            m_modelLookup = modelLookup;
            m_settings = settings;
        }

        public bool CanHandle(Type type, string prefix)
        {
            return prefix.Equals("template", StringComparison.CurrentCultureIgnoreCase) || prefix.Equals("T", StringComparison.CurrentCultureIgnoreCase);
        }

        public void ApplyFormat(FormatterContext context, Text target)
        {
            if (context.Args.Length == 0)
            {
                return;
            }
            var templateModelvariable = context.Args[0];
            if (m_modelLookup.GetValue(templateModelvariable) is not string resolvedTemplate)
            {
                throw new OpenXmlTemplateException($"Could not find template {templateModelvariable}");
            }
            Paragraph element;
            try
            {
                var templateNode = OpenXmlHelper.ParseOpenXmlString(resolvedTemplate);
                if (templateNode is Text text)
                {
                    templateNode = new Run(text);
                }
                if (templateNode is Run run)
                {
                    templateNode = new Paragraph(run);
                }
                element = templateNode as Paragraph;

            }
            catch (Exception e)
            {
                throw new OpenXmlTemplateException("Could not parse template", e);
            }
            if (element == null)
            {
                throw new OpenXmlTemplateException("Template must be a paragraph or run or text");
            }

            var templateModelLookup = new ModelLookup();
            templateModelLookup.Add("ds", context.Value);
            foreach (var models in m_modelLookup.Models.Skip(1))
            {
                templateModelLookup.Add(models.Key, models.Value);
            }
            var variableReplacer = new VariableReplacer(templateModelLookup, m_settings);
            var scriptCompiler = new ScriptCompiler(templateModelLookup, m_settings);
            var processor = new XmlNodeTemplate(element, m_settings, templateModelLookup, variableReplacer, scriptCompiler);
            processor.Process();
            var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
            var firstPart = parent.SplitAfterElement(target).First();
            firstPart.InsertAfterSelf(element);
            target.RemoveWithEmptyParent();
        }
    }
}
