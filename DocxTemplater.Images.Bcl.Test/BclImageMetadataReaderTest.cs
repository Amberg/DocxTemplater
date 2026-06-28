namespace DocxTemplater.Images.Bcl.Test
{
    internal sealed class BclImageMetadataReaderTest
    {
        private readonly BclImageMetadataReader m_reader = new();

        [Test]
        public void ReadsPngHeader()
        {
            var metadata = m_reader.Read([
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D,
                0x49, 0x48, 0x44, 0x52,
                0x00, 0x00, 0x01, 0x90,
                0x00, 0x00, 0x00, 0xF0]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Png));
            Assert.That(metadata.PixelWidth, Is.EqualTo(400));
            Assert.That(metadata.PixelHeight, Is.EqualTo(240));
        }

        [Test]
        public void ReadsGifHeader()
        {
            var metadata = m_reader.Read([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x40, 0x01, 0xF0, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Gif));
            Assert.That(metadata.PixelWidth, Is.EqualTo(320));
            Assert.That(metadata.PixelHeight, Is.EqualTo(240));
        }

        [Test]
        public void ReadsBmpHeader()
        {
            var metadata = m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x28, 0x00, 0x00, 0x00,
                0x80, 0x02, 0x00, 0x00,
                0xE0, 0x01, 0x00, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Bmp));
            Assert.That(metadata.PixelWidth, Is.EqualTo(640));
            Assert.That(metadata.PixelHeight, Is.EqualTo(480));
        }

        [Test]
        public void ReadsLittleEndianTiffHeader()
        {
            var metadata = m_reader.Read([
                0x49, 0x49, 0x2A, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x03, 0x00,
                0x00, 0x01, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x20, 0x03, 0x00, 0x00,
                0x01, 0x01, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x58, 0x02, 0x00, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Tiff));
            Assert.That(metadata.PixelWidth, Is.EqualTo(800));
            Assert.That(metadata.PixelHeight, Is.EqualTo(600));
            Assert.That(metadata.ExifRotation.Units, Is.EqualTo(180 * 60000));
        }

        [Test]
        public void ReadsJpegStartOfFrameAndExifOrientation()
        {
            var metadata = m_reader.Read([
                0xFF, 0xD8,
                0xFF, 0xE1,
                0x00, 0x22,
                0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
                0x49, 0x49, 0x2A, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0xFF, 0xC0,
                0x00, 0x11,
                0x08,
                0x01, 0x90,
                0x02, 0x58,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00,
                0xFF, 0xD9]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Jpeg));
            Assert.That(metadata.PixelWidth, Is.EqualTo(600));
            Assert.That(metadata.PixelHeight, Is.EqualTo(400));
            Assert.That(metadata.ExifRotation.Units, Is.EqualTo(90 * 60000));
        }
    }
}
