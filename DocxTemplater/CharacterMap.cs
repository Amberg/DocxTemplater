using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;



namespace DocxTemplater
{
    internal record struct Character(char Char, Text Element, int CharIndexInText);

    internal class CharacterMap
    {
        private readonly Character[] m_map;
        private readonly OpenXmlCompositeElement m_rootElement;

        public Character this[int index] => m_map[index];

        public string Text { get; private set; }

        public CharacterMap(OpenXmlCompositeElement ce)
        {
            m_rootElement = ce;
            Text = ce.InnerText;
            m_map = new Character[Text.Length];
        }

        public void Recreate()
        {
            int index = 0;
            foreach (var text in m_rootElement.Descendants<Text>())
            {
                for (var charIndexInText = 0; charIndexInText < text.Text.Length; ++charIndexInText)
                {
                    m_map[index++] = new Character(text.Text[charIndexInText], text, charIndexInText);
                }
            }
        }
    }
}