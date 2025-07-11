using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.ImageBase;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures; // http://schemas.openxmlformats.org/drawingml/2006/picture"


namespace DocxTemplater.Images
{
    public class ImageService : IImageService
    {
        private readonly Dictionary<byte[], ImageInformation> m_imagePartRelIdCache = new();
        private OpenXmlPartRootElement m_currentRoot;

        public uint GetImage(OpenXmlElement root, byte[] imageBytes, out ImageInformation imageInfoInformation)
        {
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {
                var maxPropertyId = openXmlPartRootElement.OpenXmlPart.GetMaxDocPropertyId();

                if (!TryGetImageIdFromCache(imageBytes, openXmlPartRootElement, out imageInfoInformation))
                {
                    if (IsSvgImage(imageBytes))
                    {
                        // Handle SVG specifically
                        string imagePartRelId = null;

                        if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                        {
                            imagePartRelId = CreateSvgPart(headerPart, imageBytes);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                        {
                            imagePartRelId = CreateSvgPart(footerPart, imageBytes);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                        {
                            imagePartRelId = CreateSvgPart(mainDocumentPart, imageBytes);
                        }

                        if (imagePartRelId == null)
                        {
                            throw new OpenXmlTemplateException("Could not create SVG part");
                        }

                        // Parse size arguments from SVG if possible or use defaults
                        int defaultWidth = ExtractSvgWidth(imageBytes) ?? 300;
                        int defaultHeight = ExtractSvgHeight(imageBytes) ?? 300;

                        imageInfoInformation = new ImageInformation(defaultWidth, defaultHeight, imagePartRelId, ImageRotation.CreateFromUnits(0), true);
                        m_imagePartRelIdCache[imageBytes] = imageInfoInformation;
                    }
                    else
                    {
                        using var image = Image.Load(imageBytes);
                        ImageRotation exifRotation = ImageRotation.CreateFromUnits(0);
                        if (image.Metadata?.ExifProfile?.TryGetValue(ExifTag.Orientation, out var orientationValue) == true)
                        {
                            exifRotation = ImageRotation.CreateFromExifRotation((ExifRotation)orientationValue.Value);
                        }
                        string imagePartRelId = null;
                        var imagePartType = DetectPartTypeInfo(image.Metadata);
                        if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                        {
                            imagePartRelId = CreateImagePart(headerPart, imageBytes, imagePartType);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                        {
                            imagePartRelId = CreateImagePart(footerPart, imageBytes, imagePartType);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                        {
                            imagePartRelId = CreateImagePart(mainDocumentPart, imageBytes, imagePartType);
                        }
                        if (imagePartRelId == null)
                        {
                            throw new OpenXmlTemplateException("Could not find a valid image part");
                        }
                        imageInfoInformation = new ImageInformation(image.Width, image.Height, imagePartRelId, exifRotation);
                        m_imagePartRelIdCache[imageBytes] = imageInfoInformation;
                    }
                }
                return maxPropertyId;
            }
            else
            {
                throw new OpenXmlTemplateException("Could not find root to insert image");
            }
        }

        public PIC.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, ImageRotation rotation)
        {
            return new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties
                    {
                        Id = (UInt32Value)0U,
                        Name = $"Image{propertyId}.jpg"
                    },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip(new A.BlipExtensionList(
                        new A.BlipExtension
                        {
                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                        })
                    )
                    {
                        Embed = impagepartRelationShipId,
                        CompressionState = A.BlipCompressionValues.Print
                    },
                    new A.Stretch(
                        new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = cx, Cy = cy })
                    {
                        Rotation = rotation.Units
                    },
                    new A.PresetGeometry(
                        new A.AdjustValueList()
                    )
                    {
                        Preset = A.ShapeTypeValues.Rectangle
                    }));
        }

        private bool TryGetImageIdFromCache(byte[] imageBytes, OpenXmlPartRootElement root, out ImageInformation imageInfo)
        {
            if (m_currentRoot != root)
            {
                m_imagePartRelIdCache.Clear();
                m_currentRoot = root;
            }
            return m_imagePartRelIdCache.TryGetValue(imageBytes, out imageInfo);
        }

        /// <summary>
        ///     If the image is contained in a "wsp" element (TextBox), the text box is used as a container for the image.
        ///     the size of the text box is adjusted to the size of the image.
        /// </summary>

        private static PartTypeInfo DetectPartTypeInfo(ImageMetadata imageMetadata)
        {
            return imageMetadata switch
            {
                { DecodedImageFormat.Name: "TIFF" } => ImagePartType.Tiff,
                { DecodedImageFormat.Name: "BMP" } => ImagePartType.Bmp,
                { DecodedImageFormat.Name: "GIF" } => ImagePartType.Gif,
                { DecodedImageFormat.Name: "JPEG" } => ImagePartType.Jpeg,
                { DecodedImageFormat.Name: "PNG" } => ImagePartType.Png,
                _ => throw new OpenXmlTemplateException($"Could not detect image format for image in {imageMetadata}")
            };
        }

        private static string CreateImagePart<T>(T parent, byte[] imageBytes, PartTypeInfo partType)
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(partType);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new MemoryStream(imageBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }

        private static string CreateSvgPart<T>(T parent, byte[] svgBytes)
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(ImagePartType.Svg);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new MemoryStream(svgBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }

        private static bool IsSvgImage(byte[] imageBytes)
        {
            // Check for SVG XML signature at the beginning of the file
            try
            {
                // Check the first portion of the file for SVG signature
                string content = Encoding.UTF8.GetString(imageBytes, 0, System.Math.Min(imageBytes.Length, 1000)).Trim();
                return (content.StartsWith("<?xml", System.StringComparison.OrdinalIgnoreCase) ||
                        content.StartsWith("<svg", System.StringComparison.OrdinalIgnoreCase)) &&
                       content.Contains("<svg", System.StringComparison.OrdinalIgnoreCase) &&
                       content.Contains("xmlns", System.StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static int? ExtractSvgWidth(byte[] svgBytes)
        {
            try
            {
                string svgContent = Encoding.UTF8.GetString(svgBytes);
                var widthMatch = Regex.Match(svgContent, @"<svg[^>]*\s+width\s*=\s*[""']?(\d+)(?:px)?[""']?");
                if (widthMatch.Success && int.TryParse(widthMatch.Groups[1].Value, out int width))
                {
                    return width;
                }

                // Check for viewBox as fallback
                var viewBoxMatch = Regex.Match(svgContent, @"<svg[^>]*\s+viewBox\s*=\s*[""']?(\d+)\s+(\d+)\s+(\d+)\s+(\d+)[""']?");
                if (viewBoxMatch.Success && int.TryParse(viewBoxMatch.Groups[3].Value, out int viewBoxWidth))
                {
                    return viewBoxWidth;
                }
            }
            catch
            {
                // Ignore parsing errors and use default
            }
            return null;
        }

        private static int? ExtractSvgHeight(byte[] svgBytes)
        {
            try
            {
                string svgContent = Encoding.UTF8.GetString(svgBytes);
                var heightMatch = Regex.Match(svgContent, @"<svg[^>]*\s+height\s*=\s*[""']?(\d+)(?:px)?[""']?");
                if (heightMatch.Success && int.TryParse(heightMatch.Groups[1].Value, out int height))
                {
                    return height;
                }

                // Check for viewBox as fallback
                var viewBoxMatch = Regex.Match(svgContent, @"<svg[^>]*\s+viewBox\s*=\s*[""']?(\d+)\s+(\d+)\s+(\d+)\s+(\d+)[""']?");
                if (viewBoxMatch.Success && int.TryParse(viewBoxMatch.Groups[4].Value, out int viewBoxHeight))
                {
                    return viewBoxHeight;
                }
            }
            catch
            {
                // Ignore parsing errors and use default
            }
            return null;
        }
    }
}
