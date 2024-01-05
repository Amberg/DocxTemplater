using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Templates.Formatter;
using SixLabors.ImageSharp;
using System;
using System.Linq;
using SixLabors.ImageSharp.Metadata;

namespace OpenXml.Templates.Images
{
    public class ImageFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            var prefixUpper = prefix.ToUpper();
            if (prefixUpper == "IMAGE" || prefixUpper == "IMG")
            {
                if (type == typeof(byte[]))
                {
                    return true;
                }
            }
            return false;
        }

        public void ApplyFormat(string modelPath, object value, string prefix, string[] args, Text target)
        {
            if (value is not byte[] imageBytes)
            {
                return;
            }
            try
            {
                using var image = Image.Load(imageBytes);
                var imagePartType = DetectPartTypeInfo(modelPath, image.Metadata);
                var root = target.GetRoot();
                string impagepartRelationShipId = null;
                if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
                {
                    if(openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                    {
                        impagepartRelationShipId = CreateImagePart(headerPart, imageBytes, imagePartType);
                    }
                    if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                    {
                        impagepartRelationShipId = CreateImagePart(footerPart, imageBytes, imagePartType);
                    }
                    if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                    {
                        impagepartRelationShipId = CreateImagePart(mainDocumentPart, imageBytes, imagePartType);
                    }
                }
                if(impagepartRelationShipId == null)
                {
                    throw new OpenXmlTemplateException("Could not find a valid image part");
                }

                
                var keepRatio = args.Any(x => x.ToUpper() == "KEEPRATIO");

                // case 1. Image ist the only child element of a <wps:wsp> (TextBox)
                if(TryHandleImageInWordprocessingShape(target, impagepartRelationShipId, image, keepRatio))
                {
                    return;
                }
            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new OpenXmlTemplateException("Could not detect image format", e);
            }
        }

        private int PixelToEmu(int pixel)
        {
            return (int)(pixel * 914400 / 96);
        }

        private PartTypeInfo DetectPartTypeInfo(string modelPath, ImageMetadata imageMetadata)
        {
            return imageMetadata switch
            {
                { DecodedImageFormat: { Name: "TIFF" } } => ImagePartType.Tiff,
                { DecodedImageFormat: { Name: "BMP" } } => ImagePartType.Bmp,
                { DecodedImageFormat: { Name: "GIF" } } => ImagePartType.Gif,
                { DecodedImageFormat: { Name: "JPEG" } } => ImagePartType.Jpeg,
                { DecodedImageFormat: { Name: "PNG" } } => ImagePartType.Png,
                _ => throw new OpenXmlTemplateException($"Could not detect image format for image in {modelPath}")
            };
        }

        /// <summary>
        /// If the image is contained in a <wps:wsp> element (TextBox), the text box is used as a container for the image.
        /// the size of the text box is adjusted to the size of the image.
        /// </summary>
        private bool TryHandleImageInWordprocessingShape(Text target, string impagepartRelationShipId, Image image, bool keepRatio)
        {
            var aspectRatio = image.Width / (double)image.Height;
            var shape = target.Ancestors<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape>().FirstOrDefault();
            if(shape == null)
            {
                return false;
            }
            var shapeProperties = shape.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties>();
            if (shapeProperties == null)
            {
                return false;
            }
            var blip = new DocumentFormat.OpenXml.Drawing.Blip() { Embed = impagepartRelationShipId };
            var blipFill = new DocumentFormat.OpenXml.Drawing.BlipFill(blip, new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle()))
                { Dpi = 0, RotateWithShape = true };
            shapeProperties.AddChild(blipFill);

            if (keepRatio)
            {
                // get anchor
                var anchor = shape.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
                if (anchor != null)
                {
                    var extents2 = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>();
                    // keep the with and the aspect ratio
                    extents2.Cy.Value = (long) (extents2.Cx.Value / aspectRatio);
                }
                else
                {
                    throw new OpenXmlTemplateException("Could not find anchor for shape");
                }
            }

            target.Remove();
            return true;
        }

        private string CreateImagePart<T>(T parent, byte[] imageBytes, PartTypeInfo partType) 
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(partType);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new System.IO.MemoryStream(imageBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }
    }
}
