namespace DocxTemplater.Test
{
    internal class NestedCollectionsInTableTest
    {
        [Test]
        public void NestedCollectionTest()
        {
            using var fileStream = File.OpenRead("Resources/NestedCollectionsInTable.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var date = CreateTestData();
            docTemplate.BindModel("ds", date);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        [Test]
        public void MultipleTableRowsBoundToCollection()
        {
            using var fileStream = File.OpenRead("Resources/MultipleTableRowsBoundToCollection.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var date = CreateTestData();
            docTemplate.BindModel("ds", date);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        private static object CreateTestData()
        {
            var date = new
            {
                Reports = new[]
                {
                    new
                    {
                        Date = DateTime.UtcNow,
                        TotalLessons = 9.25,
                        TotalHours = 3.0,
                        Activities = new[]
                        {
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 1",
                                Activity = "Foo",
                                Lessons = 4.0,
                            },
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 2",
                                Activity = "Bar",
                                Lessons = 3.0,
                            },
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 3",
                                Activity = "Baz",
                                Lessons = 2.0,
                            },
                        }
                    },
                    new
                    {
                        Date = DateTime.UtcNow,
                        TotalLessons = 1.75,
                        TotalHours = 0.75,
                        Activities = new[]
                        {
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 4",
                                Activity = "Qux",
                                Lessons = 1.0,
                            },
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 5",
                                Activity = "Quux",
                                Lessons = 0.5,
                            },
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 6",
                                Activity = "Quuz",
                                Lessons = 0.25,
                            },
                        }
                    },
                    new
                    {
                        Date = DateTime.UtcNow,
                        TotalLessons = 3.0,
                        TotalHours = 1.25,
                        Activities = new[]
                        {
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 7",
                                Activity = "Corge",
                                Lessons = 2.0,
                            },
                            new
                            {
                                DateAndTime = DateTime.UtcNow,
                                Description = "Activity 8",
                                Activity = "Grault",
                                Lessons = 1.0,
                            },
                        }
                    }
                }
            };
            return date;
        }
    }
}
