namespace DocxTemplater.Test
{
    internal class HtmlRenderingTest
    {

        [Test]
        public void SimpleHtmlReplacement()
        {
            //using var fileStream = File.OpenRead("Resources/SimpleHtmlRendering.docx");
            //var docTemplate = new DocxTemplate(fileStream);
            var docTemplate = DocxTemplate.Open("Resources/SimpleHtmlRendering.docx");
            var html = @"<h1>The Main Languages of the Web</h1><p>HTML is the standard markup language for creating Web pages. HTML describes the structure of a Web page, and consists of a series of elements. HTML elements tell the browser how to display the content.</p><hr><p>CSS is a language that describes how HTML elements are to be displayed on screen, paper, or in other media. CSS saves a lot of work, because it can control the layout of multiple web pages all at once.</p><hr><p>JavaScript is the programming language of HTML and the Web. JavaScript can change HTML content and attribute values. JavaScript can change CSS. JavaScript can hide and show HTML elements, and more.</p>";
            docTemplate.BindModel("ds", new { CLAUSES = html });
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }
    }
}
