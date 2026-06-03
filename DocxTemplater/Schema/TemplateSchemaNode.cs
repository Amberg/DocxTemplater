using System.Collections.Generic;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// A node in the structural schema of variables referenced by a template.
    /// Nodes form a tree rooted at <see cref="TemplateSchema.Roots"/>.
    /// </summary>
    public sealed class TemplateSchemaNode
    {
        private readonly Dictionary<string, TemplateSchemaNode> m_properties;

        internal TemplateSchemaNode(string name, TemplateNodeKind kind)
        {
            Name = name;
            Kind = kind;
            m_properties = new Dictionary<string, TemplateSchemaNode>(System.StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Name of this node as it appears in the template (case-insensitive lookups against the model).
        /// For <see cref="TemplateNodeKind.Collection"/> item schemas this is the collection's name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Kind of this node — scalar, object, or collection.
        /// </summary>
        public TemplateNodeKind Kind { get; private set; }

        /// <summary>
        /// Properties of an <see cref="TemplateNodeKind.Object"/> node (case-insensitive keys).
        /// Empty for scalars and collections.
        /// </summary>
        public IReadOnlyDictionary<string, TemplateSchemaNode> Properties => m_properties;

        /// <summary>
        /// Shape of an element for <see cref="TemplateNodeKind.Collection"/> nodes.
        /// <c>null</c> for non-collections, or when the collection's items are never accessed in the template.
        /// </summary>
        public TemplateSchemaNode ItemSchema { get; private set; }

        internal void PromoteToObject()
        {
            if (Kind == TemplateNodeKind.Scalar)
            {
                Kind = TemplateNodeKind.Object;
            }
        }

        internal void PromoteToCollection()
        {
            Kind = TemplateNodeKind.Collection;
        }

        internal TemplateSchemaNode GetOrAddProperty(string name)
        {
            PromoteToObject();
            if (!m_properties.TryGetValue(name, out var child))
            {
                child = new TemplateSchemaNode(name, TemplateNodeKind.Scalar);
                m_properties[name] = child;
            }
            return child;
        }

        internal TemplateSchemaNode EnsureItemSchema()
        {
            PromoteToCollection();
            ItemSchema ??= new TemplateSchemaNode(Name, TemplateNodeKind.Object);
            return ItemSchema;
        }
    }
}
