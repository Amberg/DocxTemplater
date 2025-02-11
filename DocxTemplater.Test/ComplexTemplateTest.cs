using DocxTemplater.Images;

namespace DocxTemplater.Test
{
    internal class ComplexTemplateTest
    {
        [Test]
        public void ProcessComplexTemplate()
        {

            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var fileStream = File.OpenRead("Resources/ComplexTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());

            var model = CreateModel(imageBytes);
            docTemplate.BindModel("ds", model);

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
        }

        [Test]
        public void ProcessComplexTemplateWithErrorHandlingInDocument()
        {
            using var fileStream = File.OpenRead("Resources/ComplexTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });

            docTemplate.BindModel("ds", new { });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            // result.SaveAsFileAndOpenInWord();

            // add items not enumerable
            fileStream.Position = 0;
            docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
            docTemplate.BindModel("ds", new
            {
                Items = new
                {
                    Images = new List<byte[]>()
                }
            });
            result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            //result.SaveAsFileAndOpenInWord();

            // add items
            fileStream.Position = 0;
            docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
            docTemplate.BindModel("ds", new { Items = new[] { new { Images = new List<byte[]>() } } });
            result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            //result.SaveAsFileAndOpenInWord();

            // add SoftwareVersions
            fileStream.Position = 0;
            docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
            docTemplate.BindModel("ds", new
            {
                Items = new[] { new
            {
                Images = new List<byte[]>(),
                SoftwareVersions = new[] {new{}}
            } }
            });
            result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            //result.SaveAsFileAndOpenInWord();

            fileStream.Position = 0;
            docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
            docTemplate.BindModel("ds", new
            {
                Items = new[] { new
                {
                    Images = new List<byte[]>(),
                    SoftwareVersions = new[] {new{ Version = "v1.0"}}
                } }
            });
            result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
        }

        private object CreateModel(byte[] imageBytes)
        {
            var items = new List<WarehouseItem>
            {
                new()
                {
                    IsHw = true,
                    Name = "Item 1",
                    HardwareRevisions = new List<VersionInfo>
                    {
                        new() {IsMajor = true, Version = "1.0"},
                        new() {IsMajor = false, Version = "1.1"},
                        new() {IsMajor = false, Version = "1.2"},
                        new() {IsMajor = false, Version = "1.3"},
                    }
                },
                new()
                {
                    IsHw = true,
                    Name = "Item 2",
                    SoftwareVersions = new List<VersionInfo>
                    {
                        new() {IsMajor = true, Version = "1.0"},
                        new() {IsMajor = false, Version = "1.1"},
                        new() {IsMajor = false, Version = "1.2"},
                        new() {IsMajor = false, Version = "1.3"},
                    }
                },
                new()
                {
                    IsHw = false,
                    Name = "Item 3",
                    SoftwareVersions = new List<VersionInfo>
                    {
                        new() {IsMajor = true, Version = "1.0"},
                        new() {IsMajor = false, Version = "1.1"},
                        new() {IsMajor = false, Version = "1.2"},
                        new() {IsMajor = false, Version = "1.3"},
                    }
                },
                new()
                {
                    IsHw = false,
                    Name = "Item 4",
                    SoftwareVersions = new List<VersionInfo>
                    {
                        new() {IsMajor = true, Version = "42.0"},
                    },
                    HardwareRevisions = new List<VersionInfo>
                    {
                        new() {IsMajor = true, Version = "1.0"},
                        new() {IsMajor = false, Version = "1.1"},
                        new() {IsMajor = false, Version = "1.2"},
                        new() {IsMajor = false, Version = "1.3"},
                    }
                }
            };

            var images = new List<byte[]>
            {
                imageBytes,
                imageBytes,
                imageBytes
            };

            return new ComplexTemplateModel
            {
                Items = items,
                Images = images
            };
        }

        private class ComplexTemplateModel
        {
            public IReadOnlyCollection<WarehouseItem> Items { get; set; }

            public IReadOnlyCollection<byte[]> Images { get; set; }
        }

        private class WarehouseItem
        {
            public bool IsHw { get; set; }

            public string Name { get; set; }

            public IReadOnlyCollection<VersionInfo> HardwareRevisions { get; set; }

            public IReadOnlyCollection<VersionInfo> SoftwareVersions { get; set; }

        }

        private class VersionInfo
        {
            public bool IsMajor { get; set; }

            public string Version { get; set; }

            public override string ToString()
            {
                return IsMajor ? $"Major Version: {Version}" : $"Minor Version: {Version}";
            }
        }
    }
}
