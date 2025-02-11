using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplater.Test
{
	internal class ShowErrorsInDocumentTest
	{

		[Test]
		public void ShowErrorsInDocument()
		{
			using var fileStream = File.OpenRead("Resources/ShowErrorsInDocument.docx");
			var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
			docTemplate.BindModel("ds", new { });
			var result = docTemplate.Process();
			docTemplate.Validate();
			result.Position = 0;
			result.SaveAsFileAndOpenInWord();
		}
	}
}
