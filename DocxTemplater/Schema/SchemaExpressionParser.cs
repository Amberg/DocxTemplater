using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using DynamicExpresso;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// Schema-only parser that recovers member-access paths from DynamicExpresso expressions.
    /// All the "reach into DynamicExpresso internals" code lives here, isolated from the rest of
    /// the codebase.
    /// </summary>
    /// <remarks>
    /// DynamicExpresso reduces dynamic dispatch into <see cref="InvocationExpression"/>s that
    /// invoke <see cref="CallSite"/>.Target. The site's binder identifies the operation
    /// (LateGetMember / LateGetIndex / LateInvokeMethod) and exposes the member or method name via
    /// a private field. We identify the binder by class name and read the field by name with
    /// defensive null-checks, so an upstream rename causes paths to be skipped rather than a crash.
    /// Tripwire tests in <c>SchemaExpressionParserTest</c> fail loudly on shape changes.
    /// </remarks>
    internal static class SchemaExpressionParser
    {

        // Pre-scan replaces leading dots (relative-scope syntax `.foo`, `..foo`) with synthetic
        // root identifiers so DynamicExpresso accepts the expression. Mirrors the leading-dot
        // rewriting in DocxTemplater.ScriptCompiler (which uses RegexWordStartingWithDot for the
        // same purpose during runtime compilation).
        private static readonly Regex LeadingDotPattern = new(
            @"(?:^|(?<unary>[+\-!])|(?<=[^.\p{L}\p{N}_?\]]))(?<dots>\.+)(?<prop>[\p{L}\p{N}_]*)",
            RegexOptions.Compiled, TimeSpan.FromMilliseconds(500));

        /// <summary>
        /// Parses <paramref name="expression"/> and emits a <see cref="TemplatePath"/> for every
        /// member-access chain. Failures during parse are silent — a broken expression will surface
        /// as an error at actual render time, not here.
        /// </summary>
        public static void Extract(string expression, Action<TemplatePath> onPath)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return;
            }

            var (rewritten, leadingDotsByRoot) = NormalizeLeadingDots(expression);

            Expression body;
            try
            {
                var interpreter = new Interpreter();
                var identifiers = interpreter.DetectIdentifiers(rewritten);
                foreach (var unknown in identifiers.UnknownIdentifiers)
                {
                    // Bind as ParameterExpression (not a constant value) so the parsed tree carries
                    // the identifier's name on its root - we attribute paths via that name.
                    interpreter.SetExpression(unknown, Expression.Parameter(typeof(SchemaPlaceholder), unknown));
                }
                body = interpreter.Parse(rewritten).Expression;
            }
            catch (Exception)
            {
                return;
            }

            var visitor = new SchemaPathVisitor(leadingDotsByRoot, onPath);
            visitor.Visit(body);
        }

        private static (string Cleaned, Dictionary<string, int> LeadingDotsByRoot) NormalizeLeadingDots(string expression)
        {
            var dotsByRoot = new Dictionary<string, int>(StringComparer.Ordinal);
            var rewritten = LeadingDotPattern.Replace(
                expression,
                m =>
                {
                    var unary = m.Groups["unary"].Value;
                    var dots = m.Groups["dots"].Value.Length;
                    var prop = m.Groups["prop"].Value;
                    if (prop.Length == 0)
                    {
                        return m.Value;
                    }
                    var synth = "__s" + dots + "_" + prop;
                    dotsByRoot[synth] = dots;
                    return unary + synth;
                });
            return (rewritten, dotsByRoot);
        }

        private sealed class SchemaPathVisitor : ExpressionVisitor
        {
            private readonly Dictionary<string, int> m_leadingDotsByRoot;
            private readonly Action<TemplatePath> m_onPath;

            public SchemaPathVisitor(Dictionary<string, int> leadingDotsByRoot, Action<TemplatePath> onPath)
            {
                m_leadingDotsByRoot = leadingDotsByRoot;
                m_onPath = onPath;
            }

            protected override Expression VisitInvocation(InvocationExpression node)
            {
                if (TryConsumeChain(node, out var path, out var subExpressionsToVisit))
                {
                    m_onPath(path);
                    foreach (var sub in subExpressionsToVisit)
                    {
                        Visit(sub);
                    }
                    return node;
                }
                return base.VisitInvocation(node);
            }

            protected override Expression VisitParameter(ParameterExpression node)
            {
                // Standalone parameter reference (body is just `.Name` after dot-rewriting, or `customer`).
                // Successful chain consumption returns without descending into children, so this fires only
                // when a parameter is referenced without any subsequent member/index/method access.
                EmitRootPath(node);
                return base.VisitParameter(node);
            }

            private void EmitRootPath(ParameterExpression param)
            {
                var rootName = param.Name ?? string.Empty;
                int dots = 0;
                if (m_leadingDotsByRoot.TryGetValue(rootName, out var d))
                {
                    dots = d;
                    rootName = StripSyntheticPrefix(rootName);
                }
                m_onPath(new TemplatePath(rootName, Array.Empty<PathSegment>()) { LeadingDotCount = dots });
            }

            private static string StripSyntheticPrefix(string syntheticName)
            {
                var underscore = syntheticName.IndexOf('_', 3);
                return underscore > 0 ? syntheticName[(underscore + 1)..] : syntheticName;
            }

            private bool TryConsumeChain(InvocationExpression node, out TemplatePath path, out List<Expression> subExpressions)
            {
                path = null;
                subExpressions = new List<Expression>();
                var segments = new List<PathSegment>();

                Expression current = node;
                while (current is InvocationExpression invoke && TryDecodeSegment(invoke, out var seg, out var target, out var subs))
                {
                    segments.Add(seg);
                    subExpressions.AddRange(subs);
                    current = target;
                }

                if (segments.Count == 0)
                {
                    return false;
                }
                if (current is not ParameterExpression param)
                {
                    return false;
                }

                segments.Reverse();

                var rootName = param.Name ?? string.Empty;
                int dots = 0;
                if (m_leadingDotsByRoot.TryGetValue(rootName, out var d))
                {
                    dots = d;
                    rootName = StripSyntheticPrefix(rootName);
                }

                path = new TemplatePath(rootName, segments) { LeadingDotCount = dots };
                return true;
            }

            private static bool TryDecodeSegment(InvocationExpression invoke, out PathSegment segment, out Expression target, out List<Expression> additionalSubExpressions)
            {
                segment = null;
                target = null;
                additionalSubExpressions = new List<Expression>();

                if (invoke.Arguments.Count < 2)
                {
                    return false;
                }
                if (invoke.Arguments[0] is not ConstantExpression callSiteConst)
                {
                    return false;
                }
                if (callSiteConst.Value is not CallSite site)
                {
                    return false;
                }

                var binderTypeName = site.Binder.GetType().Name;
                if (binderTypeName == "LateGetMemberCallSiteBinder")
                {
                    if (invoke.Arguments.Count != 2)
                    {
                        return false;
                    }
                    var name = ReadStringField(site.Binder, "_propertyOrFieldName");
                    if (name == null)
                    {
                        return false;
                    }
                    segment = new PathSegment { Kind = PathSegmentKind.Member, Name = name };
                    target = invoke.Arguments[1];
                    return true;
                }
                if (binderTypeName == "LateInvokeMethodCallSiteBinder")
                {
                    var name = ReadStringField(site.Binder, "_methodName");
                    if (name == null)
                    {
                        return false;
                    }
                    segment = new PathSegment { Kind = PathSegmentKind.Method, Name = name };
                    target = invoke.Arguments[1];
                    for (int i = 2; i < invoke.Arguments.Count; i++)
                    {
                        additionalSubExpressions.Add(invoke.Arguments[i]);
                    }
                    return true;
                }
                if (binderTypeName == "LateGetIndexCallSiteBinder")
                {
                    if (invoke.Arguments.Count < 3)
                    {
                        return false;
                    }
                    segment = new PathSegment { Kind = PathSegmentKind.Index };
                    target = invoke.Arguments[1];
                    for (int i = 2; i < invoke.Arguments.Count; i++)
                    {
                        additionalSubExpressions.Add(invoke.Arguments[i]);
                    }
                    return true;
                }
                return false;
            }

            private static string ReadStringField(CallSiteBinder binder, string fieldName)
            {
                var f = binder.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                return f?.GetValue(binder) as string;
            }
        }
    }
}
