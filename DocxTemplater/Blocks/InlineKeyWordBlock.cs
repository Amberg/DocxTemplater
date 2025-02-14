using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocxTemplater.Blocks
{
    class InlineKeyWordBlock : ContentBlock
    {
        private const string SectionBreak = "SECTIONBREAK";
        private readonly string m_keyWord;

        public InlineKeyWordBlock(IVariableReplacer variableReplacer, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(variableReplacer, patternType, startTextNode, startMatch)
        {
            m_keyWord = StartMatch.Variable.ToUpper(CultureInfo.InvariantCulture);
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            if (m_keyWord == SectionBreak)
            {
                // search paragraph of the container
                var para = StartTextNode.GetFirstAncestor<Paragraph>();
                if (para != null)
                {
                    var first = (Paragraph)para.SplitAfterElement(StartTextNode).First();
                        var breakPara = new Paragraph();
                        var paraProperties = new ParagraphProperties();
                        paraProperties.AddChild(new SectionProperties(
                                new SectionType() { Val = SectionMarkValues.NextPage }
                            ));
                            breakPara.AddChild(paraProperties);

                    first.InsertAfterSelf(breakPara);
                }

            }
            else
            {
                base.Expand(models, parentNode);
            }
        }

        public override void ExtractContentRecursively()
        {
            var res = new List<OpenXmlElement>();
            var element = m_keyWord switch
            {
                "BREAK" => (OpenXmlElement)new Break(),
                "PAGEBREAK" => new Break() {Type = BreakValues.Page},
                SectionBreak => null,
                _ => throw new OpenXmlTemplateException($"Invalid expression {StartTextNode.Text}")
            };
            res.Add(element);
            m_content = res;
        }
    }
}
