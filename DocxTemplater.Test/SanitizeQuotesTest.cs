namespace DocxTemplater.Test
{
    internal class SanitizeQuotesTest
    {
        [Test]
        public void SanitizeQuotes_SingleQuotes_ReplacedWithDoubleQuotes()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("'hello'"), Is.EqualTo("\"hello\""));
        }

        [Test]
        public void SanitizeQuotes_OpeningCurlyDoubleQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("\u201Chello"), Is.EqualTo("\"hello"));
        }

        [Test]
        public void SanitizeQuotes_ClosingCurlyDoubleQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("hello\u201D"), Is.EqualTo("hello\""));
        }

        [Test]
        public void SanitizeQuotes_LeftSingleCurlyQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("\u2018hello"), Is.EqualTo("\"hello"));
        }

        [Test]
        public void SanitizeQuotes_RightSingleCurlyQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("hello\u2019"), Is.EqualTo("hello\""));
        }

        [Test]
        public void SanitizeQuotes_EmptyString_ReturnsEmpty()
        {
            Assert.That(HelperFunctions.SanitizeQuotes(""), Is.EqualTo(""));
        }

        [Test]
        public void SanitizeQuotes_WhitespaceOnly_ReturnsEmpty()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("   "), Is.EqualTo(""));
        }

        [Test]
        public void SanitizeQuotes_NoQuotes_ReturnsInputTrimmed()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("  hello world  "), Is.EqualTo("hello world"));
        }

        [Test]
        public void SanitizeQuotes_MixedWordQuotes_AllReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("\u201Chello\u201D"), Is.EqualTo("\"hello\""));
        }

        [Test]
        public void SanitizeQuotes_LeftPointingAngleQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("\u00ABhello"), Is.EqualTo("\"hello"));
        }

        [Test]
        public void SanitizeQuotes_RightPointingAngleQuote_IsReplaced()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("hello\u00BB"), Is.EqualTo("hello\""));
        }

        [Test]
        public void SanitizeQuotes_AllQuoteTypes_AllReplacedWithDoubleQuote()
        {
            Assert.That(HelperFunctions.SanitizeQuotes("\u2018\u2019\u00AB\u00BB"), Is.EqualTo("\"\"\"\""));
        }
    }
}