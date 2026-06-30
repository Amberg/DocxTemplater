using AutoBogus;
using DocxTemplater.Images;
using DocxTemplater.Model;

namespace DocxTemplater.Test
{
    internal class MultipleRowsBoundToCollectionTest
    {
        [Test]
        public void ProcessComplexTemplate()
        {

            using var fileStream = File.OpenRead("Resources/MultipleRowsBoundToCollection.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());

            var timeReportFaker = new AutoFaker<TimeReportDate>();
            var activityFaker = new AutoFaker<TimeReportActivity>()
                .RuleFor(a => a.Duration, f => TimeSpan.FromHours(f.Random.Double(0, 4)));
            var activityTotalFaker = new AutoFaker<TimeReportActivityTotal>();
            var model = new DummyModel
            {
                FromDate = DateTime.Now.AddDays(-7),
                ToDate = DateTime.Now,
                NameFirst = "John",
                NameFamily = "Doe",
                Reports = timeReportFaker.Generate(7).ToList(),
                ActivityTotals = activityTotalFaker.Generate(3).ToList()
            };
            foreach (var report in model.Reports)
            {
                report.Activities = activityFaker.Generate(3).ToList();
            }
            docTemplate.BindModel("ds", model);
            docTemplate.BindModel("rs", new StringLocalizerDummyModel());

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
        }

        public class StringLocalizerDummyModel : ITemplateModel
        {
            public bool TryGetPropertyValue(string propertyName, out ValueWithMetadata value)
            {
                value = new ValueWithMetadata(propertyName);
                return true;
            }
        }

        private class DummyModel
        {
            public DateTime FromDate { get; set; }

            public DateTime ToDate { get; set; }

            public string NameFirst { get; set; }

            public string NameFamily { get; set; }

            public List<TimeReportDate> Reports { get; set; } = new();

            public double TotalLessons => Reports.Sum(r => r.TotalLessons);

            public double TotalHours => Reports.Sum(r => r.TotalHours);

            public List<TimeReportActivityTotal> ActivityTotals
            {
                get;
                set;
            } = new();
        }

        private class TimeReportDate
        {
            public DateTime Date { get; set; }

            public List<TimeReportActivity> Activities { get; set; } = new();

            public double TotalLessons => Activities.Sum(a => a.Lessons);

            public double TotalHours => Activities.DefaultIfEmpty().Sum(a => a.Duration.TotalHours);
        }

        private class TimeReportActivityTotal
        {
            public string Activity { get; set; }

            public double TotalLessons { get; set; }

            public double TotalHours { get; set; }
        }

        private class TimeReportActivity
        {
            public DateTime DateAndTime { get; set; }

            public double Lessons { get; set; }

            public TimeSpan Duration { get; set; }

            public string Description { get; set; }

            public string Activity { get; set; }
        }
    }
}
