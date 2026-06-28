using System.Collections.Generic;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// Structural schema of variables referenced by a template — the shape of the model
    /// a caller has to provide so the template can be rendered without missing-binding errors.
    /// </summary>
    /// <remarks>
    /// The schema is the union over all branches (both <c>if</c> and <c>else</c>, all
    /// <c>case</c>s, etc.) — a caller has to be prepared to bind everything that may be touched.
    ///
    /// Known limitations (these constructs in templates may yield an incomplete schema):
    /// <list type="bullet">
    ///   <item>Sub-template formatters (<c>:tmpl</c>): the referenced sub-template is itself a template
    ///   string at run-time and not visible to static analysis.</item>
    ///   <item>Dynamic table contents (<c>:dyntable</c>): row/column data comes from a runtime
    ///   <see cref="DocxTemplater.Model.IDynamicTable"/>; only the collection itself is reported.</item>
    ///   <item>Dictionary-style indexing with string keys (<c>props["key"]</c>) is not reflected
    ///   in the schema — the key is not statically known.</item>
    /// </list>
    /// </remarks>
    public sealed class TemplateSchema
    {
        internal TemplateSchema(IReadOnlyDictionary<string, TemplateSchemaNode> roots)
        {
            Roots = roots;
        }

        /// <summary>
        /// Root nodes of the schema, keyed by name (case-insensitive).
        /// Each entry corresponds to a top-level model that callers would pass to
        /// <see cref="TemplateProcessor.BindModel"/>.
        /// </summary>
        public IReadOnlyDictionary<string, TemplateSchemaNode> Roots { get; }
    }
}
