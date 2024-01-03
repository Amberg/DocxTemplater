using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Text;

namespace OpenXml.Templates
{
    internal record Character
    {
        public char Char;
        public OpenXmlElement Element;
        public int Index;
    }
    internal class CharacterMap
    {
        private readonly List<OpenXmlElement> m_elements = new();
        private readonly List<Character> m_map = new();
        private readonly StringBuilder m_textBuilder = new();
        private readonly OpenXmlCompositeElement m_rootElement;
        private string m_text;
        private bool m_isDirty;

        public Character this[int index]
        {
            get
            {
                if (m_isDirty)
                {
                    Recreate();
                }
                return m_map[index];
            }
        }

        public string Text
        {
            get
            {
                if (m_isDirty)
                {
                    Recreate();
                }
                return m_text;
            }
        }

        private List<OpenXmlElement> Elements
        {
            get
            {
                if (m_isDirty)
                {
                    Recreate();
                }
                return m_elements;
            }
        }

        public CharacterMap(OpenXmlCompositeElement ce)
        {
            m_rootElement = ce;
            CreateMap(m_rootElement);
            m_text = m_textBuilder.ToString();
            m_isDirty = false;
        }

        private void Recreate()
        {
            m_elements.Clear();
            m_map.Clear();
            m_textBuilder.Clear();
            CreateMap(m_rootElement);
            m_text = m_textBuilder.ToString();
            m_isDirty = false;
        }

        private void CreateMap(OpenXmlCompositeElement ce)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    CreateMap(child as OpenXmlCompositeElement);
                }
                else
                {
                    m_elements.Add(child);
                }

                if (child is Paragraph || child is Break)
                {
                    m_map.Add(new Character
                    {
                        Char = (char)10,
                        Element = child,
                        Index = -1
                    });

                    m_textBuilder.Append((char)10);
                }

                if (child is Text)
                {
                    var t = child as Text;
                    for (var i = 0; i < t.Text.Length; ++i)
                    {
                        m_map.Add(new Character
                        {
                            Char = t.Text[i],
                            Element = child,
                            Index = i
                        });
                    }

                    m_textBuilder.Append(t.Text);
                }
            }
        }
        
        public void MarkAsDirty()
        {
            m_isDirty = true;
        }
    }
}