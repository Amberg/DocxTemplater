using System.Collections.Generic;
using System.Text;

namespace DocxTemplater.Schema
{
    internal enum PathSegmentKind
    {
        /// <summary>Property/field access: <c>x.Foo</c>.</summary>
        Member,

        /// <summary>Method invocation: <c>x.Foo()</c>. Schema-wise the method is treated like a member (the return value is what's used).</summary>
        Method,

        /// <summary>Indexed access: <c>x[0]</c> or <c>x["key"]</c>.</summary>
        Index
    }

    internal sealed class PathSegment
    {
        public PathSegmentKind Kind { get; init; }

        /// <summary>Member or method name. <c>null</c> for indexed access.</summary>
        public string Name { get; init; }
    }

    /// <summary>
    /// A path through the model as referenced by a single expression: root parameter + a chain of accesses.
    /// </summary>
    internal sealed class TemplatePath
    {
        public TemplatePath(string root, IReadOnlyList<PathSegment> segments)
        {
            Root = root;
            Segments = segments;
        }

        /// <summary>Root identifier (parameter name). Empty string for paths starting with leading dots (relative scope).</summary>
        public string Root { get; }

        /// <summary>Number of leading dots in the original source. 0 for absolute paths, &gt;0 for relative-scope paths.</summary>
        public int LeadingDotCount { get; init; }

        public IReadOnlyList<PathSegment> Segments { get; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(new string('.', LeadingDotCount));
            sb.Append(Root);
            foreach (var s in Segments)
            {
                sb.Append(s.Kind switch
                {
                    PathSegmentKind.Member => "." + s.Name,
                    PathSegmentKind.Method => "." + s.Name + "()",
                    PathSegmentKind.Index => "[]",
                    _ => "?"
                });
            }
            return sb.ToString();
        }
    }
}
