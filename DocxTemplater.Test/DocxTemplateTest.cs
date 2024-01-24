using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Images;
using System.Collections;
using System.Dynamic;
using System.Globalization;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Test
{
    internal class DocxTemplateTest
    {
        [Test]
        public void DynamicTable()
        {
            using var fileStream = File.OpenRead("Resources/DynamicTable.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var tableModel = new DynamicTable();
            tableModel.AddRow(new Dictionary<object, object>() { { "Header1", "Value1" }, { "Header2", "Value2" }, { "Header3", "Value3" } });
            tableModel.AddRow(new Dictionary<object, object>() { { "Header1", "Value4" }, { "Header2", "Value5" }, { "Header3", "Value6" } });
            tableModel.AddRow(new Dictionary<object, object>() { { "Header1", "Value7" }, { "Header2", "Value8" }, { "Header3", "Value9" } });

            docTemplate.BindModel("ds", tableModel);

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var table = body.Descendants<Table>().First();
            var rows = table.Descendants<TableRow>().ToList();
            Assert.That(rows.Count, Is.EqualTo(5));
            Assert.That(rows[0].InnerText, Is.EqualTo("Header1Header2Header3"));
            Assert.That(rows[2].InnerText, Is.EqualTo("Value1Value2Value3"));
            Assert.That(rows[3].InnerText, Is.EqualTo("Value4Value5Value6"));
            Assert.That(rows[4].InnerText, Is.EqualTo("Value7Value8Value9"));
        }

        /// <summary>
        /// Dynamic tables are only required if the number of columns is not known at design time.
        /// otherwise a simple table bound to a collection of objects is sufficient.
        /// </summary>
        [Test]
        public void DynamicTableWithComplexObjectsAsHeaderAndValues()
        {
            using var fileStream = File.OpenRead("Resources/DynamicTableWithComplexObjectsAsHeaderAndValues.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.Settings.Culture = new CultureInfo("en-US");
            var tableModel = new DynamicTable();
            tableModel.AddRow(new Dictionary<object, object>()
            {
                {
                    new {HeaderTitle = "Header1"}, new { TheDouble = 20.0, TheDate = new DateTime(2007, 11, 12) }
                },
                {
                    new {HeaderTitle = "Header2"}, new { TheDouble = 30.0, TheDate = new DateTime(2007, 9, 12) }
                },
                {
                    new {HeaderTitle = "Header3"}, new { TheDouble = 40.0, TheDate = new DateTime(2001, 11, 14) }
                }
            });
            tableModel.AddRow(new Dictionary<object, object>()
            {
                {
                    new {HeaderTitle = "Header1"}, new { TheDouble = 50.0, TheDate = new DateTime(2007, 11, 12) }
                },
                {
                    new {HeaderTitle = "Header2"}, new { TheDouble = 60.0, TheDate = new DateTime(2007, 9, 12) }
                },
                {
                    new {HeaderTitle = "Header3"}, new { TheDouble = 70.0, TheDate = new DateTime(2002, 11, 9) }
                }
            });
            tableModel.AddRow(new Dictionary<object, object>()
            {
                {
                    new {HeaderTitle = "Header1"}, new { TheDouble = 80.0, TheDate = new DateTime(2007, 11, 12) }
                },
                {
                    new {HeaderTitle = "Header2"}, new { TheDouble = 90.0, TheDate = new DateTime(2007, 9, 12) }
                },
                {
                    new {HeaderTitle = "Header3"}, new { TheDouble = 100.0, TheDate = new DateTime(2003, 11, 12) }
                }
            });

            docTemplate.BindModel("ds", tableModel);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var table = body.Descendants<Table>().First();
            var rows = table.Descendants<TableRow>().ToList();
            Assert.That(rows.Count, Is.EqualTo(5));
            Assert.That(rows[0].InnerText, Is.EqualTo("HEADER1HEADER2HEADER3"));
            Assert.That(rows[2].InnerText, Is.EqualTo("20.00  11/12/200730.00  9/12/200740.00  11/14/2001"));
            Assert.That(rows[3].InnerText, Is.EqualTo("50.00  11/12/200760.00  9/12/200770.00  11/9/2002"));
            Assert.That(rows[4].InnerText, Is.EqualTo("80.00  11/12/200790.00  9/12/2007100.00  11/12/2003"));
        }

        [Test]
        public void MissingVariableThrows()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{missing}}")))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
        }

        [TestCase("<html><body><h1>Test</h1></body></html>", "<html><body><h1>Test</h1></body></html>")]
        [TestCase("<body><h1>Test</h1></body>","<html><body><h1>Test</h1></body></html>")]
        [TestCase("<h1>Test</h1>", "<html><h1>Test</h1></html>")]
        [TestCase("Test", "<html>Test</html>")]
        [TestCase("foo<br>Test", "<html>foo<br>Test</html>")]
        public void HtmlIsAlwaysEnclosedWithHtmlTags(string html, string expexted)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Here comes HTML {{ds}:html}")))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", html);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            // check word contains altChunk
            var body = document.MainDocumentPart.Document.Body;
            var altChunk = body.Descendants<AltChunk>().FirstOrDefault();
            Assert.IsNotNull(altChunk);
            // extract html part
            var htmlPart = document.MainDocumentPart.GetPartById(altChunk.Id);
            var stream = htmlPart.GetStream();
            var content = new StreamReader(stream).ReadToEnd();
            Assert.That(content, Is.EqualTo(expexted));
            // check html part contains html;
        }

        [Test]
        public void MissingVariableWithSkipErrorHandling()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Text1{{missing}}Text2{{missing2}:toupper}{{missingImg}:img()}")))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent;
            var result = docTemplate.Process();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            //check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("Text1Text2"));
        }

        [Test]
        public void LoopStartAndEndTagsAreRemoved()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Text123"))),
                new Paragraph(new Run(new Text("{{#ds.Items}}"))),
                new Paragraph(new Run(new Text("{{Items.Name}} {{Items.Price < 6}} less than 6 {{else}} more than 6{{/}}"))),
                new Paragraph(new Run(new Text("{{/ds.Items}}"))),
                new Paragraph(new Run(new Text("Text456")))
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Items = new[] { new { Name = "Item1", Price = 5 }, new { Name = "Item2", Price = 7 } } });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            // there should only be 4 paragraphs after processing
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<Paragraph>().Count(), Is.EqualTo(4));
        }

        [Test]
        public void ConditionsWithAndWithoutPrefix()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{ Test > 5 }}Test1{{ else }}else1{{ / }}"))),
                new Paragraph(new Run(new Text("{{ds.Test > 5}}Test2{{else}}else2{{/}}"))),
                new Paragraph(new Run(new Text("{{ds2.Test > 5}}Test3{{else}}else3{{/}}")))

            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Test = 6 });
            docTemplate.BindModel("ds2", new { Test = 6 });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            // check result text
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("Test1Test2Test3"));
        }

        [Test]
        public void BindToMultipleModels()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text("{{obj.var1}}")),
                new Run(new Text("{{obj.var2}}")),
                new Run(new Text("{{dynObj.var3}}")),
                new Run(new Text("{{dict.var4}}")),
                new Run(new Text("{{interface.var5}}"))
            )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);

            docTemplate.BindModel("obj", new { var1 = "var1", var2 = "var2" });
            dynamic dynObj = new ExpandoObject();
            dynObj.var3 = "var3";
            docTemplate.BindModel("dynObj", dynObj);

            var dict = new Dictionary<string, object>();
            dict.Add("var4", "var4");
            docTemplate.BindModel("dict", dict);

            var dummyModel = new DummyModel();
            dummyModel.Add("var5", "var5");
            docTemplate.BindModel("interface", dummyModel);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("var1var2var3var4var5"));
        }

        [Test]
        public void ReplaceTextBoldIsPreserved()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text("This Value:")),
                new Run(new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }), new Text("{{Property1}}")),
                new Run(new Text("will be replaced"))
            )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("Property1", "Replaced");
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open((Stream)result, false);
            var body = document.MainDocumentPart.Document.Body;
            // check that bold is preserved
            Assert.That(body.Descendants<Bold>().First().Val, Is.EqualTo(OnOffValue.FromBoolean(true)));
            // check that text is replaced
            Assert.That(body.Descendants<Text>().Skip(1).First().Text, Is.EqualTo("Replaced"));
        }

        [Test, TestCaseSource(nameof(CultureIsAppliedTest_Cases))]
        public string CultureIsAppliedTest(string formatter, CultureInfo culture, object value)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                                        new Run(new Text("Double without Formatter:")),
                                        new Run(new Text($"{{{{var}}{formatter}}}")),
                                        new Run(new Text("Double with formatter"))
                                        )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream, new ProcessSettings() { Culture = culture });
            docTemplate.BindModel("var", value);
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            return body.Descendants<Text>().Skip(1).First().Text;
        }

        static IEnumerable CultureIsAppliedTest_Cases()
        {
            yield return new TestCaseData("", new CultureInfo("en-us"), new DateTime(2024, 11, 1)).Returns("11/1/2024 12:00:00 AM");
            yield return new TestCaseData("", new CultureInfo("de-ch"), new DateTime(2024, 11, 1)).Returns("01.11.2024 00:00:00");
            yield return new TestCaseData(":f(d)", new CultureInfo("en-us"), new DateTime(2024, 11, 1, 20, 22, 33)).Returns("11/1/2024");
            yield return new TestCaseData(":FORMAT(D)", new CultureInfo("en-us"), new DateTime(2024, 11, 1, 20, 22, 33)).Returns("Friday, November 1, 2024");
            yield return new TestCaseData(":F(yyyy MM dd - HH mm ss)", new CultureInfo("en-us"), new DateTime(2024, 11, 1, 20, 22, 33)).Returns("2024 11 01 - 20 22 33");
            yield return new TestCaseData(":F(n)", new CultureInfo("en-us"), 50000.45).Returns("50,000.450");
            yield return new TestCaseData(":F(c)", new CultureInfo("en-us"), 50000.45).Returns("$50,000.45");
            yield return new TestCaseData(":F(n)", new CultureInfo("de"), 50000.45).Returns("50.000,450");
            yield return new TestCaseData(":F(c)", new CultureInfo("de-ch"), 50000.45).Returns("CHF 50’000.45");

        }


        [Test]
        public void BindCollection()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("{{PropertyInRoot}}")), // --> same as ds.PropertyInRoot
                    new Run(new Text("{{NullTest}}")),
                    new Run(new Text("{{NullTest}:TOUPPER()}")),
                    new Run(new Text("This Value:")),
                    new Run(new Text("{{#ds.Items}}"), new Text("For each run {{ds.Items.Name}}"))
                ),
            new Paragraph(
                    new Run(new Text("{{ds.Items.Value}}")),
                    new Run(new Text("{{#ds.Items.InnerCollection}}")),
                    new Run(new Text("{{Items.Value}}")), // --> same as ds.Items.Value
                    new Run(new Text("{{ds.Items.InnerCollection.Name}}")),
                    new Run(new Text("{{Items.InnerCollection.InnerValue}}")), // --> same as ds.Items.InnerCollection.InnerValue
                    new Run(new Text("{{ds.Items.InnerCollection.NumericValue > 0 }}I'm only here if NumericValue is greater than 0 - {{ds.Items.InnerCollection.InnerValue}:toupper()}{{else}}I'm here if if this is not the case{{/}}")),
                    new Run(new Text("{{/ds.Items.InnerCollection}}")),
                    new Run(new Text("{{/Items}}")), // --> same as ds.Items.InnerCollection
                    new Run(new Text("will be replaced {{company.Name}}"))
                )
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds",
                new
                {
                    PropertyInRoot = "RootValue",
                    NullTest = (string)null,
                    Items = new[]
                        {
                            new {Name = " Item1 ", Value = " Value1 ", InnerCollection = new[]
                            {
                                new {Name = " InnerItem1 ", InnerValue = " InnerValue1 ", NumericValue = 0}
                            }},
                            new {Name = " Item2 ", Value = " Value2 ", InnerCollection = new[]
                            {
                                new {Name = " InnerItem2 ", InnerValue = " InnerValue2 ", NumericValue = 0},
                                new {Name = " InnerItem2a ", InnerValue = " InnerValue2b ", NumericValue = 1}
                            }},
                        }
                });
            docTemplate.BindModel("company", new { Name = "X" });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            //check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("RootValueThis Value:For each run  Item1  Value1  Value1  InnerItem1  InnerValue1 " +
                                                   "I'm here if if this is not the caseFor each run  Item2  Value2  Value2  InnerItem2  " +
                                                   "InnerValue2 I'm here if if this is not the case Value2  InnerItem2a  InnerValue2b " +
                                                   "I'm only here if NumericValue is greater than 0 -  INNERVALUE2B will be replaced X"));
        }

        [Test]
        public void BindCollectionToTable()
        {
            var xml = @"<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">  
                      <w:tblPr>  
                        <w:tblW w:w=""5000"" w:type=""pct""/>  
                        <w:tblBorders>  
                          <w:top w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:left w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:bottom w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:right w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                        </w:tblBorders>  
                      </w:tblPr>  
                      <w:tblGrid>  
                        <w:gridCol w:w=""10296""/>  
                      </w:tblGrid>
                       <w:tr>  
                        <w:tc>  
                          <w:p><w:r><w:t>Header Col 1</w:t></w:r></w:p>  
                        </w:tc>
                        <w:tc>  
                          <w:p><w:r><w:t>Header Col 2</w:t></w:r></w:p>  
                        </w:tc>  
                      </w:tr>  
                      <w:tr>  
                        <w:tc>  
                          <w:tcPr>  
                            <w:tcW w:w=""0"" w:type=""auto""/>  
                          </w:tcPr>  
                          <w:p><w:r><w:t>{{#Items}}</w:t><w:t>{{Items.FirstVal}}</w:t></w:r></w:p>  
                        </w:tc>
                        <w:tc>  
                          <w:tcPr>  
                            <w:tcW w:w=""0"" w:type=""auto""/>  
                          </w:tcPr>  
                          <w:p><w:r><w:t>{{Items.SecondVal}}{{/Items}}</w:t></w:r></w:p>  
                        </w:tc>  
                      </w:tr>  
                    </w:tbl>";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Table(xml)));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds",
                new
                {
                    Items = new[]
                    {
                        new {FirstVal = "CC_11", SecondVal = "CC_12"},
                        new {FirstVal = "CC_21", SecondVal = "CC_22"},
                    }
                });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.IsNotNull(result);
            result.Position = 0;
            //  result.SaveAsFileAndOpenInWord();
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var table = body.Descendants<Table>().First();
            var rows = table.Descendants<TableRow>().ToList();
            Assert.That(rows.Count, Is.EqualTo(3));
            Assert.That(rows[0].InnerText, Is.EqualTo("Header Col 1Header Col 2"));
            Assert.That(rows[1].InnerText, Is.EqualTo("CC_11CC_12"));
            Assert.That(rows[2].InnerText, Is.EqualTo("CC_21CC_22"));
        }

        [Test]
        public void ProcessBillTemplate()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            using var fileStream = File.OpenRead("Resources/BillTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", new
            {
                Company = new
                {
                    Logo = imageBytes
                },
                DisplayName = "John Doe",
                Street = "Main Street 42",
                City = "New York",
                Bills = new[]
                {
                    new
                    {
                        Date = DateTime.Now,
                        Name = "Rechnung für was",
                        CustomId = "R1045",
                        Amount = 1045.5m,
                        PaidAmount = 0m,
                        OpenAmount = 1045.5m
                    },
                    new
                    {
                        Date = DateTime.Now,
                        Name = "Bill 2",
                        CustomId = "R4242",
                        Amount = 1045.5m,
                        PaidAmount = 5042m,
                        OpenAmount = 1045.5m
                    },
                },
                Total = 1045.5m,
                TotalPaid = 0m,
                TotalOpen = 1045.5m,
                TotalDownPayment = 0m,
                HtmlTest = "<br class=\"k-br\"><table class=\"k-table\"><thead><tr style=\"height:19.85pt;\">" +
                           "<th colspan=\"2\" style=\"width:538px;border-width:1px;border-style:solid;border-color:#000000;background-color:#c1bfbf;vertical-align:middle;text-align:left;margin-left:60px;" +
                           "\">Document / Notes - This is table was generated from HTML</th></tr></thead><tbody><tr style=\"height:19.85pt;\">" +
                           "<td style=\"width: 538px;\" data-role=\"resizable\">Some Notes with special characters ä ö ü é and so on</td>" +
                           "<td style=\"width:162px;text-align:left;vertical-align:top;\">29.11.2023</td></tr></tbody></table><p>&#xFEFF;</p>"
            });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Assert.That(paragraphs.Count, Is.EqualTo(61));
            // check some replacements
            Assert.That(body.InnerText.Contains("John Doe"), Is.EqualTo(true));
            Assert.That(body.InnerText.Contains("Main Street 42"), Is.EqualTo(true));
            Assert.That(body.InnerText.Contains("New York"), Is.EqualTo(true));

            // check table
            var table = body.Descendants<Table>().First();
            Assert.That(table.InnerText.Contains("Rechnung für was"), Is.EqualTo(true));
            Assert.That(table.InnerText.Contains("R1045"), Is.EqualTo(true));
            Assert.That(table.InnerText.Contains("Bill 2"), Is.EqualTo(true));
        }

        [Test]
        public void ProcessBillTemplate2()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var fileStream = File.OpenRead("Resources/BillTemplate2.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());

            var model = CrateBillTemplate2Model();
            docTemplate.BindModel("ds", model);
            docTemplate.BindModel("company", new { Logo = imageBytes });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
        }

        private static DriveStudentOverviewReportingModel CrateBillTemplate2Model()
        {
            DriveStudentOverviewReportingModel model = new()
            {
                NotBilledLessons = new List<NotBilledLessonReportModel>()
            };
            model.NotBilledLessons.Add(new NotBilledLessonReportModel()
            {
                EducationName = "Test",
                Count = 1,
                Price = 100,
                TotalPrice = 100
            });
            model.NotBilledLessons.Add(new NotBilledLessonReportModel()
            {
                EducationName = "Test2",
                Count = 2,
                Price = 200,
                TotalPrice = 400
            });
            model.Educations = new List<EducationReportModel>();
            model.Educations.Add(new EducationReportModel()
            {
                Name = "Test",
                Lessons = new List<LessonReportModel>(),
                TotalLessons = 10,
                OpenLessons = 5,
                PaidLessons = 5,
                NotBilledLessons = 2
            });
            model.Educations[0].Lessons.Add(new LessonReportModel()
            {
                Date = DateTime.Now,
                Count = 1,
                BillName = "Test",
                IsOpen = true
            });
            model.Educations[0].Lessons.Add(new LessonReportModel()
            {
                Date = DateTime.Now,
                Count = 1,
                BillName = "Test",
                IsOpen = false
            });
            model.Educations[0].Lessons.Add(new LessonReportModel()
            {
                Date = DateTime.Now,
                Count = 1,
                BillName = "Test",
                IsOpen = false
            });
            return model;
        }

        private class DriveStudentOverviewReportingModel
        {
            public List<NotBilledLessonReportModel> NotBilledLessons { get; set; }

            public List<EducationReportModel> Educations { get; set; }
        }

        private class EducationReportModel
        {
            public string Name { get; set; }

            public List<LessonReportModel> Lessons { get; set; }

            public int TotalLessons { get; set; }

            public int OpenLessons { get; set; }

            public int PaidLessons { get; set; }

            public int NotBilledLessons { get; set; }
        }

        private class LessonReportModel
        {
            public DateTime Date { get; set; }

            public double Count { get; set; }

            public string BillName { get; set; }

            public bool IsOpen
            {
                get;
                set;
            }
        }

        private class NotBilledLessonReportModel
        {
            public string EducationName { get; set; }

            public int Count { get; set; }

            public decimal Price
            {
                get;
                set;
            }

            public decimal TotalPrice
            {
                get;
                set;
            }
        }

        private class DummyModel : ITemplateModel
        {
            private readonly Dictionary<string, object> m_dict;

            public DummyModel()
            {
                m_dict = new Dictionary<string, object>();
            }

            public void Add(string key, object value)
            {
                m_dict.Add(key, value);
            }

            public bool TryGetPropertyValue(string propertyName, out object value)
            {
                return m_dict.TryGetValue(propertyName, out value);
            }
        }
    }
}
