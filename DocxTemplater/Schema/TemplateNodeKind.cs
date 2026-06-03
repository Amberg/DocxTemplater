using System.Diagnostics.CodeAnalysis;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// Kind of a node in a <see cref="TemplateSchema"/>.
    /// </summary>
    [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Schema vocabulary aligns with JSON/template terminology")]
    public enum TemplateNodeKind
    {
        /// <summary>
        /// A leaf value (string, number, bool, ...). Has no children.
        /// </summary>
        Scalar,

        /// <summary>
        /// A composite value with named properties (<see cref="TemplateSchemaNode.Properties"/>).
        /// </summary>
        Object,

        /// <summary>
        /// A collection. The shape of an element is described by <see cref="TemplateSchemaNode.ItemSchema"/>.
        /// </summary>
        Collection
    }
}
