using System;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocxTemplater.Formatter;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;
using System.Xml.Linq;
using System.Reflection;
using Path = System.IO.Path;

namespace DocxTemplater.Charts
{
    public class ChartProcessor : ITemplateProcessorExtension
    {
        private Dictionary<string, ChartReference> m_insertedChartReferences = new Dictionary<string, ChartReference>();

        public void PreProcess(OpenXmlCompositeElement content)
        {
            m_insertedChartReferences.Clear();
        }

        public void BlockContentExtracted(IReadOnlyCollection<OpenXmlElement> content)
        {
        }

        public void ReplaceVariables(IVariableReplacer variableReplacer, IModelLookup models, OpenXmlElement parentNode, List<OpenXmlElement> newContent)
        {
            // Search for chart references in the content.
            var charts = newContent.SelectMany(x => x.Descendants<ChartReference>());
            foreach (var chartReference in charts)
            {
                if (!m_insertedChartReferences.TryAdd(chartReference.Id, chartReference))
                {
                    // Get the OpenXmlPartRootElement from the chart reference.
                    var root = chartReference.GetRoot();
                    if (root is OpenXmlPartRootElement rootElement && rootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                    {
                        string newId = $"{chartReference.Id}_{m_insertedChartReferences.Count}";
                        // Get the original ChartPart using the relationship ID.
                        var originalChartPart = (ChartPart)mainDocumentPart.GetPartById(chartReference.Id);

                        // Create a new ChartPart (this will generate a new target path).
                        var clonedChartPart = mainDocumentPart.AddNewPart<ChartPart>(newId);

                        // Copy the main content of the chart.
                        using (var sourceStream = originalChartPart.GetStream(FileMode.Open))
                        using (var targetStream = clonedChartPart.GetStream(FileMode.Create))
                        {
                            sourceStream.CopyTo(targetStream);
                        }

                        // Clone the relationships from the original ChartPart.
                        CopyPartRelationshipsRecursive(originalChartPart, clonedChartPart, mainDocumentPart);


                        // Get the new relationship ID and update the chart reference.
                        var newRelId = rootElement.OpenXmlPart.GetIdOfPart(clonedChartPart);
                        chartReference.Id = newRelId;
                    }
                }
            }
        }

        /// <summary>
        /// Copies relationships from the source ChartPart to the target ChartPart.
        /// </summary>
        private void CopyPartRelationshipsRecursive(OpenXmlPart source, OpenXmlPart target, MainDocumentPart mainDocumentPart)
        {
            foreach (var rel in source.Parts)
            {
                OpenXmlPart clonedPart = null;
                if (rel.OpenXmlPart is EmbeddedPackagePart embeddedPackagePart)
                {
                    var uri = rel.OpenXmlPart.Uri;
                    // get extension
                    var extension = Path.GetExtension(uri.OriginalString);
                    clonedPart = mainDocumentPart.AddEmbeddedPackagePart(new PartTypeInfo(rel.OpenXmlPart.ContentType, extension), rel.RelationshipId);
                    var id = mainDocumentPart.GetIdOfPart(clonedPart);
                    target.AddPart(clonedPart, id);

                }
                else if (rel.OpenXmlPart is ChartColorStylePart chartColorStyle)
                {
                    clonedPart = target.AddNewPart<ChartColorStylePart>();

                }
                else if(rel.OpenXmlPart is ChartStylePart chartColorStylePart)
                {
                    clonedPart = target.AddNewPart<ChartStylePart>();
                }

                // copy content
                using (var sourceStream = rel.OpenXmlPart.GetStream(FileMode.Open))
                using (var targetStream = new MemoryStream())
                {
                    sourceStream.CopyTo(targetStream);
                    targetStream.Position = 0;  // Reset stream position before feeding data.
                    clonedPart.FeedData(targetStream);
                }

                CopyPartRelationshipsRecursive(rel.OpenXmlPart, clonedPart, mainDocumentPart);
            }
        }




    }
}