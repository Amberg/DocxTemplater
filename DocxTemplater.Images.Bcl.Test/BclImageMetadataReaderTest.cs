using System;
using System.Collections.Generic;
using NUnit.Framework;

namespace DocxTemplater.Images.Bcl.Test
{
    internal sealed class BclImageMetadataReaderTest
    {
        private readonly DefaultImageMetadataReader m_reader = new();

        private static byte[] JpegWithSegments(params byte[][] segments)
        {
            var bytes = new List<byte> { 0xFF, 0xD8 };
            foreach (var segment in segments)
            {
                bytes.AddRange(segment);
            }
            bytes.Add(0xFF);
            bytes.Add(0xD9);
            return bytes.ToArray();
        }

        private static byte[] App2Segment(ushort totalLength)
        {
            if (totalLength < 2)
            {
                throw new ArgumentOutOfRangeException(nameof(totalLength), "JPEG segment length must include the 2-byte length field.");
            }

            var payloadLength = totalLength - 2;
            var segment = new byte[2 + 2 + payloadLength];
            segment[0] = 0xFF;
            segment[1] = 0xE2;
            segment[2] = (byte)(totalLength >> 8);
            segment[3] = (byte)(totalLength & 0xFF);
            return segment;
        }

        private static byte[] JpegWithFillPaddingRun(int fillByteCount)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(fillByteCount);

            var bytes = new byte[2 + fillByteCount + 2];
            bytes[0] = 0xFF;
            bytes[1] = 0xD8;
            Array.Fill(bytes, (byte)0xFF, 2, fillByteCount);
            bytes[^2] = 0xFF;
            bytes[^1] = 0xD9;
            return bytes;
        }

        [Test]
        public void App2Segment_Throws_ForLengthsSmallerThanLengthField()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => App2Segment(1));
        }

        private static byte[] SofSegment600x400()
        {
            return [
                0xFF, 0xC0,
                0x00, 0x11,
                0x08,
                0x01, 0x90,
                0x02, 0x58,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00
            ];
        }

        private static byte[] ExifOrientationSegment(ushort orientation)
        {
            return [
                0xFF, 0xE1,
                0x00, 0x22,
                0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
                0x49, 0x49, 0x2A, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, (byte)orientation, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00
            ];
        }

        private static byte[] InvalidExifHeaderWithTiffPrefix()
        {
            return [
                0xFF, 0xE1,
                0x00, 0x22,
                0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
                0x49, 0x49, 0x2A, 0x00,
                0xFF, 0xFF, 0xFF, 0xFF,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00
            ];
        }

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
                0xE0, 0x01, 0x00, 0x00,
                0x01, 0x00, 0x18, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Bmp));
            Assert.That(metadata.PixelWidth, Is.EqualTo(640));
            Assert.That(metadata.PixelHeight, Is.EqualTo(480));
        }

        [Test]
        public void ReadsBitmapCoreBmpHeader()
        {
            var metadata = m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x0C, 0x00, 0x00, 0x00,
                0x80, 0x02,
                0xE0, 0x01,
                0x01, 0x00,
                0x18, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Bmp));
            Assert.That(metadata.PixelWidth, Is.EqualTo(640));
            Assert.That(metadata.PixelHeight, Is.EqualTo(480));
        }

        [Test]
        public void ReadsTopDownBmpHeader_AsPositivePixelHeight()
        {
            var metadata = m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x28, 0x00, 0x00, 0x00,
                0x80, 0x02, 0x00, 0x00,
                0x20, 0xFE, 0xFF, 0xFF,
                0x01, 0x00, 0x18, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Bmp));
            Assert.That(metadata.PixelWidth, Is.EqualTo(640));
            Assert.That(metadata.PixelHeight, Is.EqualTo(480));
        }

        [Test]
        public void ThrowsInvalidBmpImageHeight_WhenSignedHeightCannotBeAbsolutized()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x28, 0x00, 0x00, 0x00,
                0x80, 0x02, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x80,
                0x01, 0x00, 0x18, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00]));

            Assert.That(ex.Message, Is.EqualTo("Invalid BMP image height."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsInvalidPngIhdrChunk_ForPngWithWrongFirstChunk()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D,
                0x74, 0x45, 0x58, 0x74,
                0x00, 0x00, 0x01, 0x90,
                0x00, 0x00, 0x00, 0xF0]));

            Assert.That(ex.Message, Is.EqualTo("Invalid PNG IHDR chunk. Expected length 13 and type IHDR; found length 13 and type 0x74455874."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsInvalidPngIhdrChunk_ForPngWithWrongIhdrLength()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0C,
                0x49, 0x48, 0x44, 0x52,
                0x00, 0x00, 0x01, 0x90,
                0x00, 0x00, 0x00, 0xF0]));

            Assert.That(ex.Message, Is.EqualTo("Invalid PNG IHDR chunk. Expected length 13 and type IHDR; found length 12 and type 0x49484452."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsUnsupportedBmpDibHeader_ForUnknownBmpDibHeaderSize()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x01, 0x00, 0x00, 0x00]));

            Assert.That(ex.Message, Is.EqualTo("Unsupported BMP DIB header size 1. Supported header sizes are 12 or at least 40 bytes."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsTruncatedBmpDibHeader_WhenClaimedBmpInfoHeaderIsMissing()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x28, 0x00, 0x00, 0x00,
                0x80, 0x02, 0x00, 0x00,
                0xE0, 0x01, 0x00, 0x00]));

            Assert.That(ex.Message, Is.EqualTo("Truncated BMP DIB header. Header size 40 requires at least 54 bytes."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsTruncatedBmpDibHeader_WhenClaimedBitmapCoreHeaderIsMissing()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0x42, 0x4D,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x0C, 0x00, 0x00, 0x00,
                0x80, 0x02]));

            Assert.That(ex.Message, Is.EqualTo("Truncated BMP DIB header. Header size 12 requires at least 26 bytes."));
            Assert.That(ex.InnerException, Is.Null);
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

        [Test]
        public void ReadsJpegExifOrientation_WhenApp1ComesAfterSof()
        {
            var metadata = m_reader.Read([
                0xFF, 0xD8,
                0xFF, 0xC0,
                0x00, 0x11,
                0x08,
                0x01, 0x90,
                0x02, 0x58,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00,
                0xFF, 0xE1,
                0x00, 0x22,
                0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
                0x49, 0x49, 0x2A, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0xFF, 0xD9]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Jpeg));
            Assert.That(metadata.PixelWidth, Is.EqualTo(600));
            Assert.That(metadata.PixelHeight, Is.EqualTo(400));
            Assert.That(metadata.ExifRotation.Units, Is.EqualTo(90 * 60000));
        }

        [Test]
        public void StopsParsingOnceMetadataIsDetermined_BeforeTrailingInvalidSegments()
        {
            var metadata = m_reader.Read([
                0xFF, 0xD8,
                0xFF, 0xC0,
                0x00, 0x11,
                0x08,
                0x01, 0x90,
                0x02, 0x58,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00,
                0xFF, 0xE1,
                0x00, 0x22,
                0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
                0x49, 0x49, 0x2A, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                // This malformed segment should not be reached because parsing stops once SOF+EXIF are known.
                0xFF, 0xE2,
                0xFF, 0xFF,
                0x00,
                0xFF, 0xD9]);

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Jpeg));
            Assert.That(metadata.PixelWidth, Is.EqualTo(600));
            Assert.That(metadata.PixelHeight, Is.EqualTo(400));
            Assert.That(metadata.ExifRotation.Units, Is.EqualTo(90 * 60000));
        }

        [Test]
        public void ReadsLaterValidExif_WhenEarlierExifFailsToParse()
        {
            var metadata = m_reader.Read(JpegWithSegments(
                InvalidExifHeaderWithTiffPrefix(),
                SofSegment600x400(),
                ExifOrientationSegment(6)));

            Assert.That(metadata.Format, Is.EqualTo(ImageFormat.Jpeg));
            Assert.That(metadata.PixelWidth, Is.EqualTo(600));
            Assert.That(metadata.PixelHeight, Is.EqualTo(400));
            Assert.That(metadata.ExifRotation.Units, Is.EqualTo(90 * 60000));
        }

        [Test]
        public void ThrowsWrappedException_ForInvalidJpegSegmentLength()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read([
                0xFF, 0xD8,
                0xFF, 0xE1,
                0xFF, 0xFF,
                0x00]));

            Assert.That(ex.Message, Is.EqualTo("Invalid JPEG segment length."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsWrappedException_WhenJpegSegmentCountExceedsSafeLimit()
        {
            var segments = new byte[DefaultImageMetadataReader.MaxJpegSegmentsToScan + 1][];
            for (var i = 0; i < segments.Length; i++)
            {
                segments[i] = App2Segment(2);
            }

            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read(JpegWithSegments(segments)));

            Assert.That(ex.Message, Is.EqualTo("JPEG metadata exceeds safe parsing limits. Use an image adapter package for JPEGs with larger metadata blocks."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsWrappedException_WhenSkippedJpegSegmentPayloadBytesExceedSafeLimit()
        {
            // The metadata byte limit intentionally counts skipped segment payloads too, such as
            // large ICC profile APP2 segments before SOF, because the parser still advances across them.
            var segments = new byte[17][];
            for (var i = 0; i < segments.Length; i++)
            {
                segments[i] = App2Segment(65535);
            }

            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read(JpegWithSegments(segments)));

            Assert.That(ex.Message, Is.EqualTo("JPEG metadata exceeds safe parsing limits. Use an image adapter package for JPEGs with larger metadata blocks."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [Test]
        public void ThrowsWrappedException_WhenJpegFillPaddingRunExceedsSafeLimit()
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read(JpegWithFillPaddingRun(DefaultImageMetadataReader.MaxJpegMetadataBytesToScan + 1)));

            Assert.That(ex.Message, Is.EqualTo("JPEG metadata exceeds safe parsing limits. Use an image adapter package for JPEGs with larger metadata blocks."));
            Assert.That(ex.InnerException, Is.Null);
        }

        [TestCase(new byte[] { 0x89, 0x50, 0x4E })]
        [TestCase(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39 })]
        [TestCase(new byte[] { 0x42, 0x4D, 0x00, 0x00 })]
        [TestCase(new byte[] { 0x49, 0x49, 0x2A })]
        [TestCase(new byte[] { 0x00, 0x01, 0x02, 0x03 })]
        public void ThrowsUnsupportedOrInvalid_ForTruncatedOrUnknownSignatures(byte[] bytes)
        {
            var ex = Assert.Throws<ImageMetadataReadException>(() => m_reader.Read(bytes));

            Assert.That(ex.Message, Is.EqualTo("Unsupported or invalid image format for the default metadata reader."));
            Assert.That(ex.InnerException, Is.Null);
        }
    }
}
