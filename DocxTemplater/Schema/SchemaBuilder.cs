using System;
using System.Collections.Generic;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// Builds a <see cref="TemplateSchema"/> from a stream of path declarations emitted by
    /// the template analyzer. Mirrors the runtime resolution semantics of
    /// <see cref="DocxTemplater.ModelLookup"/>: paths whose first segment matches a loop's
    /// collection name resolve into that loop's item schema; otherwise they fall through to
    /// the next outer scope and ultimately the root.
    /// </summary>
    internal sealed class SchemaBuilder
    {
        private sealed class Frame
        {
            /// <summary>For loop scopes: dotted collection path (e.g. "items" or "customer.orders").</summary>
            public string ScopeKey { get; init; }

            /// <summary>For loop scopes: schema for an element of the collection. Paths resolved here are added below this node.</summary>
            public TemplateSchemaNode ItemSchema { get; init; }

            /// <summary>True for the root frame and for non-binding scopes (e.g. range loops).</summary>
            public bool IsRoot => ScopeKey == null;
        }

        private readonly Dictionary<string, TemplateSchemaNode> m_roots = new(StringComparer.OrdinalIgnoreCase);
        private readonly Stack<Frame> m_frames = new();

        public SchemaBuilder()
        {
            m_frames.Push(new Frame());
        }

        /// <summary>
        /// Declares <paramref name="collectionPath"/> as a collection in the appropriate scope and
        /// returns the schema node for an element of it (so the caller can open an item scope).
        /// </summary>
        public TemplateSchemaNode DeclareCollection(string collectionPath)
        {
            var path = ParseDottedPath(collectionPath);
            var node = ResolveAndDeclareLeaf(path, asCollection: true);
            return node?.EnsureItemSchema();
        }

        /// <summary>Declares a scalar path (the leaf becomes a <see cref="TemplateNodeKind.Scalar"/>).</summary>
        public void DeclareScalar(string dottedPath)
        {
            DeclareScalar(ParseDottedPath(dottedPath));
        }

        public void DeclareScalar(TemplatePath path)
        {
            ResolveAndDeclareLeaf(path, asCollection: false);
        }

        /// <summary>
        /// Pushes a loop scope. <paramref name="scopeKey"/> is the dotted collection path
        /// (matches <see cref="DocxTemplater.ModelLookup"/>'s scope-key convention).
        /// </summary>
        public void OpenItemScope(string scopeKey, TemplateSchemaNode itemSchema)
        {
            m_frames.Push(new Frame { ScopeKey = scopeKey, ItemSchema = itemSchema });
        }

        /// <summary>Pushes a transparent scope (e.g. a range loop) that introduces a scope level but no item schema.</summary>
        public void OpenTransparentScope()
        {
            m_frames.Push(new Frame { ScopeKey = "", ItemSchema = null });
        }

        public void CloseScope()
        {
            if (m_frames.Count <= 1)
            {
                throw new InvalidOperationException("Cannot close root scope");
            }
            m_frames.Pop();
        }

        public TemplateSchema Build()
        {
            return new TemplateSchema(m_roots);
        }

        /// <summary>
        /// Combines <paramref name="varName"/> (which may start with leading dots) with the enclosing
        /// scope's absolute key, so the result identifies the scope independent of where it was opened.
        /// E.g. inside loop on <c>orders</c>, <c>.lines</c> resolves to <c>orders.lines</c>.
        /// </summary>
        public string ResolveAbsoluteScopeKey(string varName)
        {
            if (string.IsNullOrEmpty(varName))
            {
                return string.Empty;
            }
            int dots = 0;
            while (dots < varName.Length && varName[dots] == '.')
            {
                dots++;
            }
            var rest = dots > 0 ? varName[dots..] : varName;
            if (dots == 0)
            {
                return rest;
            }
            // Walk up `dots` scope levels from the current top to find the outer absolute key.
            var framesArray = m_frames.ToArray(); // index 0 == top (innermost)
            if (dots - 1 < framesArray.Length)
            {
                var outer = framesArray[dots - 1].ScopeKey;
                if (!string.IsNullOrEmpty(outer))
                {
                    return outer + "." + rest;
                }
            }
            return rest;
        }

        private TemplateSchemaNode ResolveAndDeclareLeaf(TemplatePath path, bool asCollection)
        {
            if (path == null || (string.IsNullOrEmpty(path.Root) && path.Segments.Count == 0))
            {
                return null;
            }

            // For scope-key matching we only need the dotted *member* prefix (root + leading member segments
            // before any indexer/method). Indexer breaks the prefix; method ends the chain.
            var memberPrefix = BuildMemberPrefix(path);

            TemplateSchemaNode parent;
            int startSegmentIndex;
            bool rootConsumed;

            if (path.LeadingDotCount > 0)
            {
                var framesArray = m_frames.ToArray();
                int index = path.LeadingDotCount - 1;
                if (index < 0 || index >= framesArray.Length)
                {
                    return null;
                }
                var frame = framesArray[index];
                if (frame.ItemSchema == null)
                {
                    return null;
                }
                parent = frame.ItemSchema;
                startSegmentIndex = 0;
                rootConsumed = false; // the path's root is a property of the parent scope (no scope key shadows it)
            }
            else if (TryMatchScopeFrame(memberPrefix, out parent, out var matchedCount))
            {
                // matchedCount counts entries from memberPrefix that the scope key consumed. memberPrefix[0] == path.Root.
                rootConsumed = matchedCount >= 1;
                startSegmentIndex = matchedCount - 1; // segment index 0 == path.Segments[0]
                if (startSegmentIndex < 0)
                {
                    startSegmentIndex = 0;
                }
            }
            else
            {
                parent = GetOrAddRoot(path.Root);
                rootConsumed = true;
                startSegmentIndex = 0;
            }

            return Walk(parent, path, startSegmentIndex, rootConsumed, asCollection);
        }

        /// <summary>Member prefix used for scope-key matching: root + member segments up to (but not including) the first non-member step.</summary>
        private static List<string> BuildMemberPrefix(TemplatePath path)
        {
            var prefix = new List<string>(1 + path.Segments.Count);
            if (!string.IsNullOrEmpty(path.Root))
            {
                prefix.Add(path.Root);
            }
            foreach (var s in path.Segments)
            {
                if (s.Kind != PathSegmentKind.Member)
                {
                    break;
                }
                prefix.Add(s.Name);
            }
            return prefix;
        }

        /// <summary>
        /// Walks a path's segments from the resolved scope target, applying the right schema operation per segment kind:
        /// Member → descend into a property; Index → promote to Collection and descend into ItemSchema; Method → terminate.
        /// </summary>
        private static TemplateSchemaNode Walk(TemplateSchemaNode parent, TemplatePath path, int startSegmentIndex, bool rootConsumed, bool asCollection)
        {
            if (parent == null)
            {
                return null;
            }

            var current = parent;

            // If the root wasn't consumed by a scope match, treat it as a Member step from `parent`.
            if (!rootConsumed && !string.IsNullOrEmpty(path.Root))
            {
                if (IsLoopLocal(path.Root))
                {
                    return null;
                }
                current = current.GetOrAddProperty(path.Root);
            }

            for (int i = startSegmentIndex; i < path.Segments.Count; i++)
            {
                var seg = path.Segments[i];
                bool isLast = i == path.Segments.Count - 1;

                if (seg.Kind == PathSegmentKind.Method)
                {
                    // Method invocation returns a value of unknown shape. Anything after is dropped.
                    return null;
                }
                if (seg.Kind == PathSegmentKind.Index)
                {
                    current.PromoteToCollection();
                    current = current.EnsureItemSchema();
                    continue;
                }
                // Member
                if (IsLoopLocal(seg.Name))
                {
                    return null;
                }
                current = current.GetOrAddProperty(seg.Name);
                if (isLast && asCollection)
                {
                    current.PromoteToCollection();
                }
            }
            return current;
        }

        private bool TryMatchScopeFrame(List<string> memberPrefix, out TemplateSchemaNode parent, out int matchedCount)
        {
            // Iterate frames from top (innermost) to root. First scope-key that is a prefix of `memberPrefix` wins.
            foreach (var frame in m_frames)
            {
                if (frame.IsRoot || string.IsNullOrEmpty(frame.ScopeKey) || frame.ItemSchema == null)
                {
                    continue;
                }
                var keyParts = frame.ScopeKey.Split('.');
                if (keyParts.Length > memberPrefix.Count)
                {
                    continue;
                }
                bool match = true;
                for (int i = 0; i < keyParts.Length; i++)
                {
                    if (!string.Equals(keyParts[i], memberPrefix[i], StringComparison.OrdinalIgnoreCase))
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                {
                    parent = frame.ItemSchema;
                    matchedCount = keyParts.Length;
                    return true;
                }
            }
            parent = null;
            matchedCount = 0;
            return false;
        }

        private TemplateSchemaNode GetOrAddRoot(string name)
        {
            if (!m_roots.TryGetValue(name, out var node))
            {
                node = new TemplateSchemaNode(name, TemplateNodeKind.Object);
                m_roots[name] = node;
            }
            return node;
        }

        private static TemplatePath ParseDottedPath(string dottedPath)
        {
            if (string.IsNullOrEmpty(dottedPath))
            {
                return new TemplatePath(string.Empty, Array.Empty<PathSegment>());
            }
            int dots = 0;
            while (dots < dottedPath.Length && dottedPath[dots] == '.')
            {
                dots++;
            }
            var rest = dots > 0 ? dottedPath[dots..] : dottedPath;
            var parts = rest.Split('.');
            if (parts.Length == 0 || parts[0].Length == 0)
            {
                return new TemplatePath(string.Empty, Array.Empty<PathSegment>()) { LeadingDotCount = dots };
            }
            var segs = new List<PathSegment>(parts.Length - 1);
            for (int i = 1; i < parts.Length; i++)
            {
                segs.Add(new PathSegment { Kind = PathSegmentKind.Member, Name = parts[i] });
            }
            return new TemplatePath(parts[0], segs) { LeadingDotCount = dots };
        }

        private static bool IsLoopLocal(string name)
        {
            return name is "_Idx" or "_Length";
        }
    }
}
