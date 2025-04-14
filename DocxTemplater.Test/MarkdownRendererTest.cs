using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Markdown;
using System.Text;

namespace DocxTemplater.Test
{
    internal class MarkdownRendererTest
    {
        [Test]
        public void PlainText_4NewLinesCeratesEmptyParagraph()
        {
            var markdown = "Hello\r\nSecond Line\r\n\r\nSecond Paragraph First Line\r\n\r\n\r\n\r\n Line after 4 newlines --> only one paragraph\r\n\\\r\n\\\r\nLine after 2 Line breaks without paragraph\\\r\\\rLine after 2 Line breaks Linux style without paragraph";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.Descendants<Paragraph>().Count(), Is.EqualTo(3));

            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Hello</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Second Line</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Second Paragraph First Line</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Line after 4 newlines --&gt; only one paragraph</w:t></w:r><w:r><w:br /></w:r><w:r><w:br /></w:r><w:r><w:br /></w:r><w:r><w:t>Line after 2 Line breaks without paragraph</w:t></w:r><w:r><w:br /></w:r><w:r><w:br /></w:r><w:r><w:t>Line after 2 Line breaks Linux style without paragraph</w:t></w:r></w:p>"));
        }

        [Explicit]
        [Test]
        public void TestRandomString()
        {
            var markdown = "test\r\rich hbe ttest\r\rtest";
            var _ = CreateTemplateWithMarkdownAndReturnBody(markdown);
        }

        [Test]
        public void ListAfterNewline()
        {
            var markdown = "**Krisenmanagement**\r- Schulung Führungsunterstützung\r- Schulung Krisenstab";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:rPr><w:b /></w:rPr><w:t>Krisenmanagement</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Schulung Führungsunterstützung</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Schulung Krisenstab</w:t></w:r></w:p>"));
        }


        [Test]
        public void ListAfterHeading()
        {
            string markdown = """
                              * First
                              * Second
                              """;
            using var fileStream = File.OpenRead("Resources/ListTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", new Dictionary<string, object>() { { "MyMarkdown", new ValueWithMetadata(markdown, new ValueMetadata("md")) } });

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            Assert.That(TestHelper.ComputeSha256Hash(document.MainDocumentPart.Document.Body.InnerXml), Is.EqualTo("09ea0bc7feaf0322847cc90a18594dea401adec7f0c7e22f2a9c1b2ec2ef1891"));
        }

        [Test]
        public void TestListRenderer()
        {
            string markdown = """
                                             - Die generelle Zielsetzung kann wie folgt umschrieben werden:  
                                             - Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025
                                            
                                             Text
                                             
                                             - Die generelle Zielsetzung kann wie folgt umschrieben werden:  
                                             - Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025
                                             """;
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);

            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Die generelle Zielsetzung kann wie folgt umschrieben werden:</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Text</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Die generelle Zielsetzung kann wie folgt umschrieben werden:</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025</w:t></w:r></w:p>"));
        }

        [Test]
        public void GridTableRenderer()
        {
            string markdown = """
                        +-----+:---:+----:+
                        |  A  |  B  |  C  |
                        +-----+-----+-----+
                        |  1  |  2  |  3  |
                        +-----+-----+-----+
                        """;

            using var fileStream = File.OpenRead("Resources/MarkdownTableCopiesStyleFromExistingTable.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", new Dictionary<string, object>() { { "MyMarkdown", new ValueWithMetadata(markdown, new ValueMetadata("md")) } });

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            Assert.That(TestHelper.ComputeSha256Hash(document.MainDocumentPart.Document.Body.InnerXml), Is.EqualTo("11a698a4f6eb4afb33d2815810f6160d29b4fe065156a2c5886a845af3ce5a55"));
        }

        [Test]
        public void MarkdownRenderDefinedInMetadata()
        {
            string markdown = """
                        - Die generelle Zielsetzung kann wie folgt umschrieben werden
                        - Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025
                        
                        | Documents / Meetings | Date              |
                        | --- | --------------------------------|
                        | Risk Analysis Region X - Scenario Description and Assessment | 29.03.2010 |
                        | PLAN-X Guide - Regional Hazard Analysis and Preparedness | 01.01.2013 |
                        | Meeting between A. Sample (Dept. A) and B. Example (Dept. B) | 16.02.2023 |
                        | Meeting between A. Sample (Dept. A) and B. Example (Dept. B) | 15.03.2023 |
                        
                        Line 1
                        - Die generelle Zielsetzung kann wie folgt umschrieben werden:  
                        - Schulung/Training der Führungsunterstützung mit hohem Praxisbezug, auf Basis der Erkenntnisse aus der Krisenstabsübung 2025
                        
                        
                        
                        
                        Line 2
                        
                        | Documents / Meetings | Date |
                        | --- | ---:|
                        | Email request from C. Example | 18.07.2024 |
                        | On-site meeting between D. Testman and C. Example (both from Clinic X) as well as E. Sample and F. Demo (both from Authority Y) | 25.09.2024 |
                        | Received documents: Crisis Manual Clinic X, Safety Guidelines 2024, Evacuation Plan Clinic X | 01.10.2024 |
                        
                        Line After table
                        """;

            using var fileStream = File.OpenRead("Resources/MarkdownTableCopiesStyleFromExistingTable.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", new Dictionary<string, object>() { { "MyMarkdown", new ValueWithMetadata(markdown, new ValueMetadata("md")) } });

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
        }


        [Test]
        public void TestTwoListsWithoutGap()
        {
            string markdown = """
                              * First
                              * Second
                              
                              - A
                              - B
                              """;
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>"));
        }

        [Test]
        public void TestTwoTablesWithoutGap()
        {
            string markdown = """
                              +-----+:---:+----:+
                              |  A  |  B  |  C  |
                              +-----+-----+-----+
                              |  1  |  2  |  3  |
                              +-----+-----+-----+
                              
                              +-----+:---:+----:+
                              |  D  |  E  |  F  |
                              +-----+-----+-----+
                              |  4  |  5  |  6  |
                              +-----+-----+-----+
                              
                              First List
                              * First
                              * Second
                              
                              Second List
                              * A
                              * B
                              """;

            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\" />" +
                                                  "</w:tblPr><w:tblGrid><w:gridCol /><w:gridCol /><w:gridCol /></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" />" +
                                                  "</w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\" />" +
                                                  "</w:pPr><w:r><w:t>B</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" />" +
                                                  "</w:pPr><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:r><w:t>1</w:t>" +
                                                  "</w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\" /></w:pPr><w:r><w:t>2</w:t>" +
                                                  "</w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" /></w:pPr><w:r><w:t>3</w:t>" +
                                                  "</w:r></w:p></w:tc></w:tr></w:tbl><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />" +
                                                  "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\" /></w:tblPr><w:tblGrid>" +
                                                  "<w:gridCol /><w:gridCol /><w:gridCol /></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:r><w:t>D</w:t></w:r>" +
                                                  "</w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\" /></w:pPr><w:r><w:t>E</w:t></w:r></w:p>" +
                                                  "</w:tc><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" /></w:pPr><w:r><w:t>F</w:t></w:r></w:p></w:tc>" +
                                                  "</w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc><w:tc><w:tcPr>" +
                                                  "<w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\" /></w:pPr><w:r><w:t>5</w:t></w:r></w:p></w:tc><w:tc><w:tcPr>" +
                                                  "<w:tcW w:w=\"1650\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" /></w:pPr><w:r><w:t>6</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>First List</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" />" +
                                                  "<w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" />" +
                                                  "<w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Second List</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" />" +
                                                  "<w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" />" +
                                                  "<w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>"));
        }

        [Test]
        public void DifferentTableStyleDefinedWithInlineSettings()
        {
            string markdown = """
                              | Documents / Meetings | Date |
                              | --- | ---:|
                              | Risk Analysis Region X - Scenario Description and Assessment | 29.03.2010 |
                              | PLAN-X Guide - Regional Hazard Analysis and Preparedness | 01.01.2013 |
                              | Meeting between A. Sample (Dept. A) and B. Example (Dept. B) | 16.02.2023 |
                              | Meeting between A. Sample (Dept. A) and B. Example (Dept. B) | 15.03.2023 |
                              """;

            using var fileStream = File.OpenRead("Resources/MarkdownTablesDifferentStyle.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new DocxTemplater.Markdown.MarkdownFormatter());
            docTemplate.BindModel("ds", markdown);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            Assert.That(TestHelper.ComputeSha256Hash(document.MainDocumentPart.Document.Body.InnerXml), Is.EqualTo("83fbf2ba90daddbb0d4fd1ae92b5e98f2c6f4f8a686f75b8813cf02a95d59fbe"));

        }


        [Test]
        public void MarkdownWithPlaceholderReplacement()
        {
            var markdown = "_Hello_ **{{ds.Name}:ToUpper}** Doe";
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document =
                new Document(new Body(new Paragraph(
                    new Run(new Text("Line Before ")),
                    new Run(new Text("{{ds.markdown}:md}")),
                    new Run(new Text("Line After"))
                    )));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", new
            {
                Name = "John",
                markdown
            });

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;

            // {{ds.markdown}:md} --> "_Hello_ **{{ds:Name}}**" --> "Hello John"
            Assert.That(body.InnerText, Is.EqualTo("Line Before Hello JOHN DoeLine After"));
            Assert.That(body.Descendants<Paragraph>().Count(), Is.EqualTo(1));
        }

        [TestCase("~~Hello~~")]
        public void XXText(string markdown)
        {
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var runWithHello = (Run)body.Descendants<Text>().Single(x => x.Text == "Hello").Parent;
            Assert.That(runWithHello?.RunProperties?.Strike, Is.Not.Null);
            Assert.That(runWithHello.RunProperties.Italic, Is.Null);
        }

        [TestCase("**Hello**")]
        [TestCase("__Hello__")]
        public void BoldText(string markdown)
        {
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var runWithHello = (Run)body.Descendants<Text>().Single(x => x.Text == "Hello").Parent;
            Assert.That(runWithHello.RunProperties.Bold, Is.Not.Null);
            Assert.That(runWithHello.RunProperties.Italic, Is.Null);
        }

        [TestCase("*Hello*")]
        [TestCase("_Hello_")]
        public void ItalicText(string markdown)
        {
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var runWithHello = (Run)body.Descendants<Text>().Single(x => x.Text == "Hello").Parent;
            Assert.That(runWithHello.RunProperties.Bold, Is.Null);
            Assert.That(runWithHello.RunProperties.Italic, Is.Not.Null);
        }

        [Test]
        public void MixedText()
        {
            var body = CreateTemplateWithMarkdownAndReturnBody("_Hello_ **There**");
            var runWithHello = (Run)body.Descendants<Text>().Single(x => x.Text == "Hello").Parent;
            var runWithThere = (Run)body.Descendants<Text>().Single(x => x.Text == "There").Parent;
            Assert.That(runWithHello.RunProperties.Bold, Is.Null);
            Assert.That(runWithHello.RunProperties.Italic, Is.Not.Null);
            Assert.That(runWithThere.RunProperties.Bold, Is.Not.Null);
            Assert.That(runWithThere.RunProperties.Italic, Is.Null);
            Assert.That(body.InnerText, Is.EqualTo("Hello There"));
        }


        [TestCase("***Hello***")]
        [TestCase("___Hello___")]
        [TestCase("**_Hello_**")]
        [TestCase("__*Hello*__")]
        public void BoldAndItalic(string markdown)
        {
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var runWithHello = (Run)body.Descendants<Text>().Single(x => x.Text == "Hello").Parent;
            Assert.That(runWithHello.RunProperties.Bold, Is.Not.Null);
            Assert.That(runWithHello.RunProperties.Italic, Is.Not.Null);
        }

        [TestCase("> Dorothy followed her through many of the beautiful rooms in her castle.\r\n>\r\n>> The Witch bade her clean the pots and kettles and sweep the floor and keep the fire fed with wood.")]
        public void BlockQuote_JustIgnored(string markdown)
        {
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Dorothy" +
                                                  " followed her through many of the beautiful rooms in her castle.</w:t></w:r></w:p><w:p xmlns:w=" +
                                                  "\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>The Witch bade her clean" +
                                                  " the pots and kettles and sweep the floor and keep the fire fed with wood.</w:t></w:r></w:p>"));
        }

        [Test]
        public void MultipleLines()
        {
            var sb = new StringBuilder();
            sb.AppendLine("__This is bold__");
            sb.AppendLine("This is not");
            sb.AppendLine("**This is bold**");
            sb.AppendLine("This is not");
            sb.AppendLine("_This is italic_");
            sb.AppendLine("This is not");
            sb.AppendLine("*This is italic*");
            sb.AppendLine("This is not");
            sb.AppendLine("***This is bold and italic***");
            var body = CreateTemplateWithMarkdownAndReturnBody(sb.ToString());
            var lineCount = body.Descendants<Run>().Count(x => x.ChildElements.Any(x => x is Break));
            Assert.That(lineCount, Is.EqualTo(8));
        }


        [Test]
        public void OrderedList()
        {
            var markdown = "1. First\n2. Second\n3. Third\n 4. Fourth\n 5. Fifth";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Third</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Fourth</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Fifth</w:t></w:r></w:p>"));
        }

        [Test]
        public void UnorderedList()
        {
            var markdown = "* First\n* Second\n* Third";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Third</w:t></w:r></w:p>"));
        }

        [Test]
        public void NestedUnorderedList()
        {
            var sb = new StringBuilder();
            sb.AppendLine("* First");
            sb.AppendLine("* Second");
            sb.AppendLine("  * First First");
            sb.AppendLine("  * First Second");
            sb.AppendLine("    * First Second First");
            sb.AppendLine("    * First Second Second");
            sb.AppendLine("  * First Third");
            sb.AppendLine("* Third");

            var body = CreateTemplateWithMarkdownAndReturnBody(sb.ToString());
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"2\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"2\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Third</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Third</w:t></w:r></w:p>"));
        }

        [Test]
        public void NestedOrderedList()
        {
            var sb = new StringBuilder();
            sb.AppendLine("1. First");
            sb.AppendLine("2. Second");
            sb.AppendLine("   1. First First");
            sb.AppendLine("   2. First Second");
            sb.AppendLine("      1. First Second First");
            sb.AppendLine("      2. First Second Second");
            sb.AppendLine("   3. First Third");
            sb.AppendLine("3. Third");
            var body = CreateTemplateWithMarkdownAndReturnBody(sb.ToString());
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"2\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second First</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"2\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Second Second</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"1\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>First Third</w:t></w:r></w:p>" +
                                                  "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"ListParagraph\" /><w:numPr><w:ilvl w:val=\"0\" /><w:numId w:val=\"1\" /></w:numPr></w:pPr><w:r><w:t>Third</w:t></w:r></w:p>"));
        }

        [Test]
        public void Table()
        {
            var sb = new StringBuilder();
            sb.AppendLine("| Header 1 | Header 2 |");
            sb.AppendLine("|:----------|----------:|");
            sb.AppendLine("| Row 1 Col 1 | Row 1 Col 2 |");
            sb.AppendLine("| Row 2 Col 1 | Row 2 Col 2 |");
            var body = CreateTemplateWithMarkdownAndReturnBody(sb.ToString());
            Assert.That(body.InnerXml, Is.EqualTo("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\" />" +
                                                  "</w:tblPr><w:tblGrid><w:gridCol /><w:gridCol /><w:gridCol /></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\" />" +
                                                  "</w:tcPr><w:p><w:pPr><w:jc w:val=\"left\" /></w:pPr><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\" />" +
                                                  "</w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" /></w:pPr><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr>" +
                                                  "<w:tcW w:w=\"2500\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"left\" /></w:pPr><w:r><w:t>Row 1 Col 1</w:t></w:r></w:p></w:tc><w:tc>" +
                                                  "<w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" /></w:pPr><w:r><w:t>Row 1 Col 2</w:t></w:r></w:p>" +
                                                  "</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"left\" />" +
                                                  "</w:pPr><w:r><w:t>Row 2 Col 1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\" /></w:tcPr><w:p><w:pPr><w:jc w:val=\"right\" />" +
                                                  "</w:pPr><w:r><w:t>Row 2 Col 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />"));
        }

        [Test]
        public void Headings()
        {

            var sb = new StringBuilder();
            sb.AppendLine("# Heading 1");
            sb.AppendLine("Some text after heading");
            sb.AppendLine("Next Line After heading");
            sb.AppendLine("## Heading 2");
            sb.AppendLine("Some text after heading");
            sb.AppendLine("Next Line After heading");
            sb.AppendLine("### Heading 3");
            sb.AppendLine("Some text after heading");
            sb.AppendLine("Next Line After heading");
            sb.AppendLine("#### Heading 4");
            sb.AppendLine("Some text after heading");
            sb.AppendLine("Next Line After heading");
            sb.AppendLine("##### Heading 5");
            sb.AppendLine("Some text after heading");
            sb.AppendLine("Next Line After heading");
            sb.AppendLine("###### Heading 6");
            var body = CreateTemplateWithMarkdownAndReturnBody(sb.ToString());
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading1\" /></w:pPr><w:r><w:t>Heading 1</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Some text after heading</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Next Line After heading</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading2\" /></w:pPr><w:r><w:t>Heading 2</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Some text after heading</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Next Line After heading</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading3\" /></w:pPr><w:r><w:t>Heading 3</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Some text after heading</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Next Line After heading</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading4\" /></w:pPr><w:r><w:t>Heading 4</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Some text after heading</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Next Line After heading</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading5\" /></w:pPr><w:r><w:t>Heading 5</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Some text after heading</w:t></w:r><w:r><w:br /></w:r><w:r><w:t>Next Line After heading</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:pStyle w:val=\"Heading6\" /></w:pPr><w:r><w:t>Heading 6</w:t></w:r></w:p><w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />"));
        }

        [Test]
        public void TestMarkDownInTemplateWithEmbeddedStyles()
        {
            using var fileStream = File.OpenRead("Resources/MarkdownTestTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream);

            var sb = new StringBuilder();
            sb.AppendLine("# This is a Table in a Table");
            sb.AppendLine("Created from markdown");
            sb.AppendLine();
            sb.AppendLine("| Header 1 | Header 2 |");
            sb.AppendLine("|----------|----------|");
            sb.AppendLine("| Row 1 Col 1 | Row 1 Col 2 |");
            sb.AppendLine("| Row 2 Col 1 | _Row 2_ **Col 2** |");

            var sb2 = new StringBuilder();
            sb2.AppendLine("# This is a List in a Table");
            sb2.AppendLine("Created from markdown");
            sb2.AppendLine("1. First");
            sb2.AppendLine("2. Second");
            sb2.AppendLine("   1. First First");
            sb2.AppendLine("   2. First Second");
            sb2.AppendLine("      1. First Second First");
            sb2.AppendLine("      2. First Second Second");
            sb2.AppendLine("   3. First Third");
            sb2.AppendLine("3. Third");

            var model = new
            {
                Numbers = new int[] { 1, 2, 3, 4, 5 },
                markdown = File.ReadAllText("Resources/TestMarkDown.md"),
                markdownInTable = sb.ToString(),
                markdownInTableList = sb2.ToString(),
                footer = "_Footer_",
                header = "**Header**",
                bold = "**Bold**",
            };
            docTemplate.RegisterFormatter(new DocxTemplater.Markdown.MarkdownFormatter(new MarkDownFormatterConfiguration()
            {
                TableStyle = "MarkDownTestTableStyle"
            }));
            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
        }

        [Test]
        public void CustomListStyleInTemplate()
        {
            using var fileStream = File.OpenRead("Resources/CustomMarkdownListStyleInTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var sb = new StringBuilder();
            sb.AppendLine("* First");
            sb.AppendLine("* Second");
            sb.AppendLine("  * First First");
            sb.AppendLine("  * First Second");
            sb.AppendLine("    * First Second First");
            sb.AppendLine("    * First Second Second");
            sb.AppendLine("  * First Third");
            sb.AppendLine("* Third");
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", sb.ToString());
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
        }

        private Body CreateTemplateWithMarkdownAndReturnBody(string markdown)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document =
                new Document(new Body(new Paragraph(new Run(new Text("{{ds}:md}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new DocxTemplater.Markdown.MarkdownFormatter());
            docTemplate.BindModel("ds", markdown);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body;
        }
    }
}
