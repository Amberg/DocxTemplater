

using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.Model;

namespace DocxTemplater.Test
{
    class TimeTrackingTemplateTest
    {
        [Test]
        public void ProcessDocument()
        {
            using var fileStream = File.OpenRead("Resources/TimeTracking.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { Culture = new CultureInfo("de-CH") });
            var data = new TimeTrackingReportPageModel
            {
                TotalDuration = TimeSpan.FromHours(10),
                TotalBillableDuration = TimeSpan.FromHours(5),
                Items = new List<TimeTrackingItemModel>
                {
                    new()
                    {
                        Date = new DateTime(2025,2,25),
                        Duration = TimeSpan.FromHours(1),
                        BillableDuration = TimeSpan.FromHours(0.5),
                        Activity = "Activity 1",
                        Description = "Description 1",
                        Owner = "Owner 1",
                        Institution = "Institution 1",
                        Person = "Person 1",
                        State = "State 1",
                        Title = "Title 1",
                    },
                    new()
                    {
                        Date = new DateTime(2025,2,26),
                        Duration = TimeSpan.FromHours(2),
                        BillableDuration = TimeSpan.FromHours(1),
                        Activity = "Activity 2",
                        Description = "Description 2",
                        Owner = "Owner 2",
                        Institution = "Institution 2",
                        Person = "Person 2",
                        State = "State 2",
                        Title = "Title 2",
                    }
                }
            };

            docTemplate.BindModel("ds", data);
            docTemplate.BindModel("gus", new { ProjectTitle = "Test Project" });

            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("Test ProjectDatumAktivitätFunktionStundengesamtVerechnete Stunden25.02.2025Title 1 Description 1Activity 11.00 h0.50 h 26.02.2025Title 2 Description 2Activity 22.00 h1.00 h Total10 h5 h"));
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
        }

        internal class TimeTrackingReportPageModel
        {
            [ModelProperty(DefaultFormatter = "F(g)")]
            public TimeSpan TotalDuration { get; set; }

            [ModelProperty(DefaultFormatter = "F(g)")]
            public TimeSpan TotalBillableDuration { get; set; }

            public List<TimeTrackingItemModel> Items { get; set; } = new List<TimeTrackingItemModel>();
        }

        internal class TimeTrackingItemModel
        {
            [ModelProperty(DefaultFormatter = "F(g)")]
            public TimeSpan Duration { get; set; }

            [ModelProperty(DefaultFormatter = "F(F2)")]
            public double DurationHours => Duration.TotalHours;

            [ModelProperty(DefaultFormatter = "F(g)")]
            public TimeSpan BillableDuration { get; set; }

            [ModelProperty(DefaultFormatter = "F(F2)")]
            public double BillableDurationHours => BillableDuration.TotalHours;

            [ModelProperty(DefaultFormatter = "F(d)")]
            public DateTime Date { get; set; }

            public string Activity { get; set; }

            public string Description { get; set; }

            public string Owner { get; set; }

            public string Institution { get; set; }

            public string Person { get; set; }

            public string State { get; set; }

            public string Title
            {
                get;
                set;
            }
        }
    }
}
