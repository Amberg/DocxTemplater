using DocumentFormat.OpenXml.Packaging;
using Markdig.Syntax.Inlines;
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    internal sealed class ImageInlineRenderer : OpenXmlObjectRenderer<LinkInline>
    {
        private readonly MainDocumentPart m_mainDocumentPart;
        private readonly IImageService m_imageService;

        public ImageInlineRenderer(MainDocumentPart mainDocumentPart, IImageService imageService)
        {
            m_mainDocumentPart = mainDocumentPart;
            m_imageService = imageService;
        }

        protected override void Write(MarkdownToOpenXmlRenderer renderer, LinkInline obj)
        {
            if (!obj.IsImage || string.IsNullOrEmpty(obj.Url) || !obj.Url.StartsWith("data:image"))
            {
                return;
            }

            var root = m_mainDocumentPart.RootElement;
            byte[] imageBytes;

            try
            {
                imageBytes = Convert.FromBase64String(obj.Url.Split(',')[1]);
            }
            catch (Exception ex)
            {
                throw new OpenXmlTemplateException($"Invalid image data in Markdown link. {obj.Url}", ex);
            }

            var maxPropertyId = m_imageService.GetImage(root, imageBytes, out ImageInformation imageInfo);
            var drawing = CreateDrawing(imageInfo, maxPropertyId, m_imageService);
            var existingText = renderer.CurrentParagraph.Descendants<Text>().LastOrDefault();
            if (existingText != null)
            {
                existingText.InsertAfterSelf(drawing);
            }
            else
            {
                renderer.CurrentParagraph.Append(new Run(drawing));
            }
        }

        private static Drawing CreateDrawing(ImageInformation imageInfo, uint maxDocumentPropertyId, IImageService imageService)
        {
            var propertyId = maxDocumentPropertyId + 1;

            return
                new Drawing(
                    new DW.Inline(
                        new DW.Extent
                        {
                            Cx = OpenXmlHelper.PixelsToEmu(imageInfo.PixelWidth),
                            Cy = OpenXmlHelper.PixelsToEmu(imageInfo.PixelHeight)
                        },
                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties
                        {
                            Id = propertyId,
                            Name = $"Picture {propertyId}",
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                    imageService.CreatePicture(imageInfo.ImagePartRelationId, propertyId, OpenXmlHelper.PixelsToEmu(imageInfo.PixelWidth), OpenXmlHelper.PixelsToEmu(imageInfo.PixelHeight), imageInfo.ExifRotation)
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U
                    });
        }
    }
}
