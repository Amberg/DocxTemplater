using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;



namespace DocxTemplater
{
    internal record struct CharacterPointer(Text Element, int CharIndexInText, int Index);

    internal class CharacterMap
    {
        private readonly CharacterPointer[] m_map;
        private readonly OpenXmlCompositeElement m_rootElement;
        private readonly StringBuilder m_stringBuilder;

        public CharacterPointer this[int index] => m_map[index];

        public string Text { get; private set; }

        public CharacterMap(OpenXmlCompositeElement ce)
        {
            m_rootElement = ce;
            Text = ce.InnerText;
            m_map = new CharacterPointer[Text.Length];
            m_stringBuilder = new StringBuilder(Text.Length);
            Recreate();
        }

        public void Recreate()
        {
            int index = 0;
            m_stringBuilder.Clear();
            foreach (var text in m_rootElement.Descendants<Text>())
            {
                m_stringBuilder.Append(text.Text);
                for (var charIndexInText = 0; charIndexInText < text.Text.Length; ++charIndexInText)
                {
                    m_map[index] = new CharacterPointer(text, charIndexInText, index);
                    ++index;
                }
            }

            Text = m_stringBuilder.ToString();
        }


        public CharacterPointer MergeText(CharacterPointer first, CharacterPointer last)
        {
            var commonParent = first.Element.FindCommonParent(last.Element) ?? throw new ArgumentException("Text elements must have a common parent");
            if (first.CharIndexInText < 0 || first.CharIndexInText >= first.Element.Text.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(first));
            }
            var matchLength = last.Index - first.Index + 1;
            var matchInSameElement = first.Element == last.Element;
            if (first.CharIndexInText != 0) // leading text is not part of the match
            {
                // new last part
                var newPart = first.Element.SplitAtIndexAndReturnLastPart(first.CharIndexInText);
                var index = first.Index;
                InsertNewTextAddIndex(index, newPart);
                first = m_map[index];
            }
            if (first.CharIndexInText + matchLength < first.Element.Text.Length) // trailing text part of the match in the first element
            {
                var newPart = first.Element.SplitAtIndexAndReturnLastPart(first.CharIndexInText + matchLength);
                InsertNewTextAddIndex(first.Index + matchLength, newPart);
            }

            if (matchInSameElement)
            {
                return first;
            }

            List<OpenXmlElement> toRemove = new();
            bool found = false;
            foreach (var current in commonParent.Descendants<Text>())
            {
                if (current == first.Element)
                {
                    found = true;
                    continue;
                }
                if (!found)
                {
                    continue;
                }

                // trailing text is not part of the match int the last element
                if (first.Element.Text.Length + current.Text.Length > matchLength)
                {
                    var firstPart = current.Text[..(matchLength - first.Element.Text.Length)];
                    var secondPart = current.Text[(matchLength - first.Element.Text.Length)..];
                    AddTextToCharacterPointer(first, firstPart);
                    current.Text = secondPart;
                    FixMapAtIndex(first.Index + first.Element.Text.Length);
                    break;
                }
                AddTextToCharacterPointer(first, current.Text);
                toRemove.Add(current);
                if (current == last.Element)
                {
                    break;
                }
            }
            foreach (var text in toRemove)
            {
                text.RemoveWithEmptyParent();
            }
            return first;
        }

        private void FixMapAtIndex(int firstIndex)
        {
            if (firstIndex < m_map.Length)
            {
                var start = m_map[firstIndex];
                if (start.Element != m_map[firstIndex - 1].Element)
                {
                    InsertNewTextAddIndex(firstIndex, start.Element);
                }
            }

        }

        private void InsertNewTextAddIndex(int startIndex, Text newPart)
        {
            for (int i = 0; i < newPart.Text.Length; i++)
            {
                var index = startIndex + i;
                m_map[index] = new CharacterPointer(newPart, i, index);
            }
        }

        private void AddTextToCharacterPointer(CharacterPointer characterPointer, string appendText)
        {
            var textStart = characterPointer.Index - characterPointer.CharIndexInText;
            for (int i = 0; i < appendText.Length; i++)
            {
                var offset = characterPointer.Element.Text.Length + i;
                var index = textStart + offset;
                m_map[index] = new CharacterPointer(characterPointer.Element, offset, index);
            }
            characterPointer.Element.Text += appendText;
        }
    }
}