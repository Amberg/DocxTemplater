using System;
using System.Buffers.Binary;
using DocxTemplater.ImageBase;

namespace DocxTemplater.Images.Bcl
{
    /// <summary>
    /// Reads image dimensions, format and EXIF orientation using only .NET base class library APIs.
    /// </summary>
    /// <remarks>
    /// This adapter deliberately reads file headers only. It does not decode, transform, resize or validate the full image payload.
    /// </remarks>
    public sealed class BclImageMetadataReader : IImageMetadataReader
    {
        private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        private static readonly byte[] Gif87aSignature = [0x47, 0x49, 0x46, 0x38, 0x37, 0x61];
        private static readonly byte[] Gif89aSignature = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61];

        public ImageMetadata Read(byte[] imageBytes)
        {
            if (imageBytes == null)
            {
                throw new ArgumentNullException(nameof(imageBytes));
            }

            try
            {
                var bytes = imageBytes.AsSpan();
                if (TryReadPng(bytes, out var metadata)
                    || TryReadGif(bytes, out metadata)
                    || TryReadBmp(bytes, out metadata)
                    || TryReadTiff(bytes, out metadata)
                    || TryReadJpeg(bytes, out metadata))
                {
                    return metadata;
                }
            }
            catch (Exception e) when (e is ArgumentOutOfRangeException or OverflowException)
            {
                throw new ImageMetadataReadException("Could not read image metadata using the .NET BCL adapter.", e);
            }

            throw new ImageMetadataReadException("Unsupported or invalid image format for the .NET BCL adapter.", null);
        }

        private static bool TryReadPng(ReadOnlySpan<byte> bytes, out ImageMetadata metadata)
        {
            metadata = null;
            if (bytes.Length < 24 || !bytes[..PngSignature.Length].SequenceEqual(PngSignature))
            {
                return false;
            }

            var width = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(16, 4)));
            var height = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(20, 4)));
            metadata = CreateMetadata(width, height, ImageFormat.Png);
            return true;
        }

        private static bool TryReadGif(ReadOnlySpan<byte> bytes, out ImageMetadata metadata)
        {
            metadata = null;
            if (bytes.Length < 10
                || (!bytes[..Gif87aSignature.Length].SequenceEqual(Gif87aSignature)
                    && !bytes[..Gif89aSignature.Length].SequenceEqual(Gif89aSignature)))
            {
                return false;
            }

            var width = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6, 2));
            var height = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8, 2));
            metadata = CreateMetadata(width, height, ImageFormat.Gif);
            return true;
        }

        private static bool TryReadBmp(ReadOnlySpan<byte> bytes, out ImageMetadata metadata)
        {
            metadata = null;
            if (bytes.Length < 26 || bytes[0] != 0x42 || bytes[1] != 0x4D)
            {
                return false;
            }

            var width = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(18, 4));
            var signedHeight = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(22, 4));
            metadata = CreateMetadata(width, Math.Abs(signedHeight), ImageFormat.Bmp);
            return true;
        }

        private static bool TryReadJpeg(ReadOnlySpan<byte> bytes, out ImageMetadata metadata)
        {
            metadata = null;
            if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
            {
                return false;
            }

            var position = 2;
            var width = 0;
            var height = 0;
            var rotation = ImageRotation.CreateFromUnits(0);

            while (position < bytes.Length)
            {
                while (position < bytes.Length && bytes[position] == 0xFF)
                {
                    position++;
                }

                if (position >= bytes.Length)
                {
                    break;
                }

                var marker = bytes[position++];
                if (marker is 0xD9 or 0xDA)
                {
                    break;
                }

                if (marker is 0x01 or >= 0xD0 and <= 0xD7)
                {
                    continue;
                }

                if (position + 2 > bytes.Length)
                {
                    break;
                }

                var segmentLength = BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(position, 2));
                position += 2;
                if (segmentLength < 2 || position + segmentLength - 2 > bytes.Length)
                {
                    throw new ArgumentOutOfRangeException(nameof(bytes), "Invalid JPEG segment length.");
                }

                var segment = bytes.Slice(position, segmentLength - 2);
                if (marker == 0xE1)
                {
                    rotation = ReadExifRotationFromApplicationSegment(segment);
                }
                else if (IsStartOfFrameMarker(marker) && segment.Length >= 7)
                {
                    height = BinaryPrimitives.ReadUInt16BigEndian(segment.Slice(1, 2));
                    width = BinaryPrimitives.ReadUInt16BigEndian(segment.Slice(3, 2));
                }

                position += segment.Length;
            }

            if (width <= 0 || height <= 0)
            {
                throw new ImageMetadataReadException("Could not find a JPEG start-of-frame segment.", null);
            }

            metadata = new ImageMetadata(width, height, ImageFormat.Jpeg, rotation);
            return true;
        }

        private static bool TryReadTiff(ReadOnlySpan<byte> bytes, out ImageMetadata metadata)
        {
            metadata = null;
            if (!TryReadTiffHeader(bytes, out var littleEndian, out var firstIfdOffset))
            {
                return false;
            }

            if (!TryReadTiffDirectory(bytes, firstIfdOffset, littleEndian, out var width, out var height, out var rotation)
                || width == null
                || height == null)
            {
                throw new ImageMetadataReadException("Could not find required TIFF width and height tags.", null);
            }

            metadata = new ImageMetadata(width.Value, height.Value, ImageFormat.Tiff, rotation);
            return true;
        }

        private static bool TryReadTiffHeader(ReadOnlySpan<byte> bytes, out bool littleEndian, out int firstIfdOffset)
        {
            littleEndian = false;
            firstIfdOffset = 0;
            if (bytes.Length < 8)
            {
                return false;
            }

            if (bytes[0] == 0x49 && bytes[1] == 0x49)
            {
                littleEndian = true;
            }
            else if (bytes[0] == 0x4D && bytes[1] == 0x4D)
            {
                littleEndian = false;
            }
            else
            {
                return false;
            }

            if (ReadUInt16(bytes.Slice(2, 2), littleEndian) != 42)
            {
                return false;
            }

            firstIfdOffset = checked((int)ReadUInt32(bytes.Slice(4, 4), littleEndian));
            return true;
        }

        private static bool TryReadTiffDirectory(
            ReadOnlySpan<byte> bytes,
            int directoryOffset,
            bool littleEndian,
            out int? width,
            out int? height,
            out ImageRotation rotation)
        {
            width = null;
            height = null;
            rotation = ImageRotation.CreateFromUnits(0);

            if (directoryOffset < 0 || directoryOffset + 2 > bytes.Length)
            {
                return false;
            }

            var entryCount = ReadUInt16(bytes.Slice(directoryOffset, 2), littleEndian);
            var entriesStart = directoryOffset + 2;
            if (entryCount > (bytes.Length - entriesStart) / 12)
            {
                return false;
            }

            for (var i = 0; i < entryCount; i++)
            {
                var entry = bytes.Slice(entriesStart + (i * 12), 12);
                var tag = ReadUInt16(entry[..2], littleEndian);
                var type = ReadUInt16(entry.Slice(2, 2), littleEndian);
                var count = ReadUInt32(entry.Slice(4, 4), littleEndian);
                if (count != 1)
                {
                    continue;
                }

                if ((tag == 0x0100 || tag == 0x0101) && TryReadTiffScalar(entry, type, littleEndian, out var scalar))
                {
                    if (tag == 0x0100)
                    {
                        width = checked((int)scalar);
                    }
                    else
                    {
                        height = checked((int)scalar);
                    }
                }
                else if (tag == 0x0112 && type == 3)
                {
                    rotation = CreateRotationFromExifOrientation(ReadUInt16(entry.Slice(8, 2), littleEndian));
                }
            }

            return true;
        }

        private static bool TryReadTiffScalar(ReadOnlySpan<byte> entry, ushort type, bool littleEndian, out uint value)
        {
            switch (type)
            {
                case 3:
                    value = ReadUInt16(entry.Slice(8, 2), littleEndian);
                    return true;
                case 4:
                    value = ReadUInt32(entry.Slice(8, 4), littleEndian);
                    return true;
                default:
                    value = 0;
                    return false;
            }
        }

        private static ImageRotation ReadExifRotationFromApplicationSegment(ReadOnlySpan<byte> segment)
        {
            if (segment.Length < 14
                || segment[0] != 0x45
                || segment[1] != 0x78
                || segment[2] != 0x69
                || segment[3] != 0x66
                || segment[4] != 0x00
                || segment[5] != 0x00)
            {
                return ImageRotation.CreateFromUnits(0);
            }

            var tiff = segment[6..];
            if (!TryReadTiffHeader(tiff, out var littleEndian, out var firstIfdOffset)
                || !TryReadTiffDirectory(tiff, firstIfdOffset, littleEndian, out _, out _, out var rotation))
            {
                return ImageRotation.CreateFromUnits(0);
            }

            return rotation;
        }

        private static ImageRotation CreateRotationFromExifOrientation(ushort orientation)
        {
            return orientation switch
            {
                3 or 6 or 8 => ImageRotation.CreateFromExifRotation((ExifRotation)orientation),
                _ => ImageRotation.CreateFromUnits(0)
            };
        }

        private static bool IsStartOfFrameMarker(byte marker)
        {
            return marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
        }

        private static ImageMetadata CreateMetadata(int width, int height, ImageFormat format)
        {
            if (width <= 0 || height <= 0)
            {
                throw new ImageMetadataReadException($"Invalid image dimensions {width}x{height}.", null);
            }

            return new ImageMetadata(width, height, format, ImageRotation.CreateFromUnits(0));
        }

        private static uint ReadUInt32(ReadOnlySpan<byte> bytes, bool littleEndian)
        {
            return littleEndian
                ? BinaryPrimitives.ReadUInt32LittleEndian(bytes)
                : BinaryPrimitives.ReadUInt32BigEndian(bytes);
        }

        private static ushort ReadUInt16(ReadOnlySpan<byte> bytes, bool littleEndian)
        {
            return littleEndian
                ? BinaryPrimitives.ReadUInt16LittleEndian(bytes)
                : BinaryPrimitives.ReadUInt16BigEndian(bytes);
        }
    }
}
