using System;
using System.Collections;
using System.Globalization;
using System.Linq;

namespace DocxTemplater
{
    /// <summary>
    /// The tag and the iteration count of a range loop <c>{{@i:count}}</c>, shared by the renderers of all template types.
    /// </summary>
    internal static class RangeLoopCount
    {
        /// <summary>The name of the index variable when the tag does not name one (<c>{{@3}}</c>).</summary>
        public const string DefaultIndexName = "Index";

        /// <summary>
        /// Splits the variable of a range loop tag into the name of the index variable and the count:
        /// <c>i:3</c> gives (<c>i</c>, <c>3</c>); without a colon the index is named <see cref="DefaultIndexName"/>.
        /// </summary>
        public static (string IndexName, string Count) Parse(string variable)
        {
            variable ??= string.Empty;
            var colon = variable.IndexOf(':');
            return colon < 0
                ? (DefaultIndexName, variable.Trim())
                : (variable[..colon].Trim(), variable[(colon + 1)..].Trim());
        }

        /// <summary>
        /// Resolves the count: an integer literal, or a model value that is an integer, a string holding an integer
        /// or a collection, whose element count is used. A value that is <c>null</c> counts as 0.
        /// </summary>
        /// <exception cref="OpenXmlTemplateException">
        /// The value cannot be resolved (a binding error), or it is neither numeric nor enumerable
        /// (<see cref="TemplateErrorCode.RangeCountNotNumeric"/>, <see cref="TemplateErrorCode.RangeCountNotNumericOrEnumerable"/>).
        /// </exception>
        public static int Resolve(IModelLookup models, string count, ProcessSettings settings)
        {
            if (int.TryParse(count, NumberStyles.Integer, CultureInfo.InvariantCulture, out int literal))
            {
                return literal;
            }

            var model = models.GetValue(count);
            switch (model)
            {
                case null:
                    return 0;
                case string text:
                    if (int.TryParse(text, out int parsed))
                    {
                        return parsed;
                    }
                    throw OpenXmlTemplateException.Create(settings, TemplateErrorCode.RangeCountNotNumeric, count, text);
            }

            try
            {
                return Convert.ToInt32(model);
            }
            catch (Exception e) when (e is InvalidCastException or FormatException or OverflowException)
            {
                if (model is IEnumerable enumerable)
                {
                    var items = enumerable.Cast<object>();
                    return items.TryGetNonEnumeratedCount(out int itemCount) ? itemCount : items.Count();
                }
                throw OpenXmlTemplateException.Create(settings, e, TemplateErrorCode.RangeCountNotNumericOrEnumerable, count, model.GetType().FullName);
            }
        }
    }
}
