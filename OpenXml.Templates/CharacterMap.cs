using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenXml.Templates
{
    internal record Character
    {
        public char Char;
        public OpenXmlElement Element;
        public int Index;
    }

    internal record MapPart
    {
        public int StartIndex;
        public int EndIndex;
    }


    internal class CharacterMap
    {
        private static readonly Regex m_regex = new(@"\{\{([a-zA-Z0-9\.]+)\}(?::(\w+\(*\w*\)*))*\}", RegexOptions.Compiled);
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

        public int GetIndex(Text text)
        {
            // Can be used to get the index of a CodeBlock.Placeholder.
            // Then you can replace text that occurs after the code block only (instead of all text).
            var index = Elements.IndexOf(text);
            if (index == -1)
            {
                return -1;
            }

            for (var i = index; i >= 0; --i)
            {
                var t = Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return m_map.IndexOf(m_map.First(c => c.Element == t && c.Index == t.Text.Length - 1));
                }
            }

            for (var i = index + 1; i < Elements.Count; ++i)
            {
                var t = Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return m_map.IndexOf(m_map.First(c => c.Element == t && c.Index == 0));
                }
            }

            return 0;
        }

        internal void ReplaceText(string oldValue, string newValue, int startIndex = 0,
            StringComparison stringComparison = StringComparison.InvariantCulture)
        {
            var i = Text.IndexOf(oldValue, startIndex, stringComparison);
            var dirty = i != -1;
            while (i != -1)
            {
                var part = new MapPart
                {
                    StartIndex = i,
                    EndIndex = i + oldValue.Length - 1
                };

                ReplaceText(part, newValue);

                startIndex = i + newValue.Length;
                i = Text.IndexOf(oldValue, startIndex, stringComparison);
            }
            m_isDirty = dirty;
        }

        internal void ReplaceVariables(ModelDictionary models)
        {
            var matches = ((IReadOnlyCollection<Match>)m_regex.Matches(Text)).Select(m => new
            {
                VariableName = m.Groups[1].Value,
                Text = (Text)this[m.Index].Element,
                MatchIndex = m.Index,
                Length = m.Length,
                Format = m.Groups[2].Value
            }).ToList();
            foreach (var m in matches)
            {
                var value = models.GetValue(m.VariableName);
                ReplaceTextAtIndex(m.MatchIndex, m.Length, value?.ToString());
            }
        }

        internal void ReplaceTextAtIndex(int startIndex, int length, string newValue)
        {
            var part = new MapPart
            {
                StartIndex = startIndex,
                EndIndex = startIndex + length - 1
            };
            ReplaceText(part, newValue);
        }

        private void ReplaceText(MapPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            var parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (var i = Elements.IndexOf(endText); i >= Elements.IndexOf(startText); --i)
            {
                var element = Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }

                if (element == startText)
                {
                    startText.Space = SpaceProcessingModeValues.Preserve;

                    if (startText == endText)
                    {
                        string postScriptum = null;
                        if (endIndex + 1 != startText.Text.Length)
                        {
                            postScriptum = startText.Text.Substring(endIndex + 1);
                        }

                        startText.Text = startText.Text.Substring(0, startIndex);
                        if (!string.IsNullOrEmpty(newText))
                        {
                            startText.Text += newText;
                        }
                        startText.Text += postScriptum;
                    }
                    else
                    {
                        startText.Text = startText.Text.Substring(0, startIndex);
                        if (!string.IsNullOrEmpty(newText))
                        {
                            startText.Text += newText;
                        }
                    }
                }
                else if (element == endText)
                {
                    endText.Space = SpaceProcessingModeValues.Preserve;
                    endText.Text = endText.Text.Substring(endIndex + 1);
                }
                else
                {
                    element.Remove();
                }
            }
        }

        internal Text ReplaceWithText(MapPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            var addedText = new Text
            {
                Text = newText,
                Space = SpaceProcessingModeValues.Preserve
            };

            var parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (var i = Elements.IndexOf(endText); i >= Elements.IndexOf(startText); --i)
            {
                var element = Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }
                if (element == startText)
                {
                    startText.Space = SpaceProcessingModeValues.Preserve;
                    if (startText == endText)
                    {
                        string postScriptum = null;
                        if (endIndex + 1 != startText.Text.Length)
                        {
                            postScriptum = startText.Text.Substring(endIndex + 1);
                        }

                        startText.Text = startText.Text.Substring(0, startIndex);

                        if (!string.IsNullOrEmpty(postScriptum))
                        {
                            startText.InsertAfterSelf(new Text
                            {
                                Text = postScriptum,
                                Space = SpaceProcessingModeValues.Preserve
                            });
                        }

                        startText.InsertAfterSelf(addedText);
                    }
                    else
                    {
                        startText.Text = startText.Text.Substring(0, startIndex);
                        startText.InsertAfterSelf(addedText);
                    }
                }
                else if (element == endText)
                {
                    endText.Space = SpaceProcessingModeValues.Preserve;
                    endText.Text = endText.Text.Substring(endIndex + 1);
                }
                else
                {
                    element.Remove();
                }
            }

            return addedText;
        }

       
        public void MarkAsDirty()
        {
            m_isDirty = true;
        }
    }
}