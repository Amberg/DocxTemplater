using DocxTemplater.Test.Contracts;
using NUnit.Framework;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;

namespace DocxTemplater.Images.ImageSharp.Test
{
    internal sealed class ImageSharpImageFormatterContractTests : ImageFormatterAdapterContractTests
    {
        protected override IEnumerable<ImageRasterCase> RasterCases =>
            [
                new ImageRasterCase("jpg", "testImage"),
                new ImageRasterCase("tiff", "testImage"),
                new ImageRasterCase("png", "testImage"),
                new ImageRasterCase("png", "testImage_rot"),
                new ImageRasterCase("bmp", "testImage"),
                new ImageRasterCase("gif", "testImage")
            ];

        protected override DocxTemplater.Formatter.IFormatter CreateFormatter()
        {
            return new ImageSharpImageFormatter();
        }

        protected override byte[] ConvertToFormat(byte[] sourceImageBytes, string extension)
        {
            using var img = Image.Load(sourceImageBytes);
            Assert.That(img.Configuration.ImageFormatsManager.TryFindFormatByFileExtension(extension, out var format));
            var memStream = new MemoryStream();
            img.Save(memStream, format);
            return memStream.ToArray();
        }

        protected override byte[] GetLargeRasterImageBytes()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var img = Image.Load(imageBytes);
            img.Mutate(x => x.Resize(img.Width * 10, img.Height * 10));

            using var bigImgStream = new MemoryStream();
            img.SaveAsJpeg(bigImgStream);
            return bigImgStream.ToArray();
        }
    }

    internal sealed class ImageSharpMetadataReaderContractTests : ImageMetadataReaderAdapterContractTests
    {
        protected override IEnumerable<ImageRasterCase> RasterCases =>
            [
                new ImageRasterCase("jpg", "testImage"),
                new ImageRasterCase("tiff", "testImage"),
                new ImageRasterCase("png", "testImage"),
                new ImageRasterCase("bmp", "testImage"),
                new ImageRasterCase("gif", "testImage")
            ];

        protected override IImageMetadataReader CreateReader()
        {
            return new ImageSharpImageMetadataReader();
        }

        protected override byte[] ConvertToFormat(byte[] sourceImageBytes, string extension)
        {
            using var img = Image.Load(sourceImageBytes);
            Assert.That(img.Configuration.ImageFormatsManager.TryFindFormatByFileExtension(extension, out var format));
            using var memStream = new MemoryStream();
            img.Save(memStream, format);
            return memStream.ToArray();
        }
    }
}
