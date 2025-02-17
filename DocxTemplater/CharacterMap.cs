using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;



namespace DocxTemplater
{
    internal record struct Character(char Char, Text Element, int CharIndexInText);

    internal class CharacterMap
    {
        private readonly Character[] m_map;
        private readonly OpenXmlCompositeElement m_rootElement;
        private readonly StringBuilder m_stringBuilder;

        public Character this[int index] => m_map[index];

        public string Text { get; private set; }

        public CharacterMap(OpenXmlCompositeElement ce)
        {
            m_rootElement = ce;
            Text = ce.InnerText;
            m_map = new Character[Text.Length];
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
                    m_map[index++] = new Character(text.Text[charIndexInText], text, charIndexInText);
                }
            }

            Text = m_stringBuilder.ToString();
        }
    }
}