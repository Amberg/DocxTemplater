using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures; // http://schemas.openxmlformats.org/drawingml/2006/picture"


namespace DocxTemplater.Images
{
    public class ImageService : IImageService
    {
        private readonly Dictionary<byte[], ImageInformation> m_imagePartRelIdCache = new();
        private OpenXmlPartRootElement m_currentRoot;

        public ImageService(IImageMetadataReader imageMetadataReader)
        {
            ImageMetadataReader = imageMetadataReader ?? throw new ArgumentNullException(nameof(imageMetadataReader));
        }

        internal IImageMetadataReader ImageMetadataReader { get; }

        public uint GetImage(OpenXmlElement root, byte[] imageBytes, out ImageInformation imageInfoInformation)
        {
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {
                var maxPropertyId = openXmlPartRootElement.OpenXmlPart.GetMaxDocPropertyId();

                if (!TryGetImageIdFromCache(imageBytes, openXmlPartRootElement, out imageInfoInformation))
                {
                    if (SvgHelper.TryReadAsSvg(imageBytes, out int width, out int height))
                    {
                        // Handle SVG specifically. It can be embedded without a bitmap decoder.
                        string imagePartRelId = null;
                        if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                        {
                            imagePartRelId = SvgHelper.CreateSvgPart(headerPart, imageBytes);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                        {
                            imagePartRelId = SvgHelper.CreateSvgPart(footerPart, imageBytes);
                        }
                        else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                        {
                            imagePartRelId = SvgHelper.CreateSvgPart(mainDocumentPart, imageBytes);
                        }

                        if (imagePartRelId == null)
                        {
                            throw new OpenXmlTemplateException("Could not create SVG part");
                        }

                        imageInfoInformation = new ImageInformation(width, height, imagePartRelId, ImageRotation.CreateFromUnits(0), true);
                        m_imagePartRelIdCache[imageBytes] = imageInfoInformation;
                    }
                    else
                    {
                        var metadata = ImageMetadataReader.Read(imageBytes);
                        string imagePartRelId = null;
                        var imagePartType = DetectPartTypeInfo(metadata.Format);
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
                        imageInfoInformation = new ImageInformation(metadata.PixelWidth, metadata.PixelHeight, imagePartRelId, metadata.ExifRotation);
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

        private static PartTypeInfo DetectPartTypeInfo(ImageFormat imageFormat)
        {
            return imageFormat switch
            {
                ImageFormat.Tiff => ImagePartType.Tiff,
                ImageFormat.Bmp => ImagePartType.Bmp,
                ImageFormat.Gif => ImagePartType.Gif,
                ImageFormat.Jpeg => ImagePartType.Jpeg,
                ImageFormat.Png => ImagePartType.Png,
                _ => throw new OpenXmlTemplateException($"Could not detect image format for image format {imageFormat}")
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
    }
}
