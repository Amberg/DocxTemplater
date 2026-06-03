using System.Dynamic;

namespace DocxTemplater.Schema
{
    /// <summary>
    /// Tracking placeholder bound to each unknown identifier when the schema extractor parses
    /// an expression with DynamicExpresso. Being a <see cref="DynamicObject"/> forces
    /// DynamicExpresso to emit dynamic dispatch — a CallSite-based tree the extractor introspects
    /// to recover member access paths. The placeholder itself is never executed; only its presence
    /// at parse time matters.
    /// </summary>
    internal sealed class SchemaPlaceholder : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = new SchemaPlaceholder();
            return true;
        }

        public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
        {
            result = new SchemaPlaceholder();
            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            result = new SchemaPlaceholder();
            return true;
        }
    }
}
