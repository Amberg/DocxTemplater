// Compatibility shims for older frameworks (netstandard2.0 / net48)
#if NETSTANDARD2_0 || NET48
namespace System.Runtime.CompilerServices
{
    // Provide IsExternalInit type for init-only support in older frameworks
    internal static class IsExternalInit { }
}
#endif
