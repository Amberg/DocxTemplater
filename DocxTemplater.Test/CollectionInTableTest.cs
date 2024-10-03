namespace DocxTemplater.Test
{
    internal class CollectionInTableTest
    {
        [Test]
        public void CollectionInTableTestRender()
        {

            using var fileStream = File.OpenRead("Resources/CollectionInTableCell.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var data = new MyLessonsReportModel();
            var lesson1 = new LessonReportModel
            {
                Date = DateTime.Now,
                CourseDisplayName = "Course 1",
                ParticipantsCount = 10,
                Resources = new List<string> { "Resource 1", "Resource 2" }
            };
            var lesson2 = new LessonReportModel
            {
                Date = DateTime.Now.AddDays(1),
                CourseDisplayName = "Course 2",
                ParticipantsCount = 20,
                Resources = new List<string> { "Resource 3", "Resource 4" }
            };
            data.Lessons = new List<LessonReportModel> { lesson1 };
            docTemplate.BindModel("ds", data);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        public class MyLessonsReportModel
        {
            public IReadOnlyCollection<LessonReportModel> Lessons { get; set; } = new List<LessonReportModel>();
        }

        public class LessonReportModel
        {
            public DateTime Date { get; set; }

            public string CourseDisplayName { get; set; }

            public ICollection<string> Resources { get; set; } = new List<string>();

            public int ParticipantsCount { get; set; }
        }
    }
}