using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// German texts for the error messages of DocxTemplater.
    /// Register them once at startup: <c>TemplateErrorMessages.Default.AddGerman();</c>
    /// </summary>
    public static class GermanErrorMessages
    {
        /// <summary>The neutral culture the texts are registered for; <c>de-CH</c>, <c>de-AT</c>, ... fall back to it.</summary>
        public static CultureInfo Culture { get; } = new("de");

        /// <summary>
        /// Adds the German texts to <paramref name="messages"/>.
        /// </summary>
        /// <returns><paramref name="messages"/>, to chain calls.</returns>
        public static TemplateErrorMessages AddGerman(this TemplateErrorMessages messages)
        {
            return messages.AddLanguage(Culture, Formats);
        }

        public static IReadOnlyDictionary<TemplateErrorCode, string> Formats { get; } = new ReadOnlyDictionary<TemplateErrorCode, string>(
            new Dictionary<TemplateErrorCode, string>
            {
                // model binding
                [TemplateErrorCode.ModelPrefixAlreadyBound] = "Ein Modell mit dem Präfix '{0}' wurde bereits gebunden - Modell-Präfixe unterscheiden nicht zwischen Groß- und Kleinschreibung",
                [TemplateErrorCode.ModelNotFound] = "Modell {0} nicht gefunden",
                [TemplateErrorCode.PropertyNotFound] = "Eigenschaft {0} nicht gefunden in {1}",
                [TemplateErrorCode.PropertyNotFoundInParentScope] = "Eigenschaft {0} nicht gefunden im übergeordneten Bereich",
                [TemplateErrorCode.PropertyNotFoundOnCollection] = "Eigenschaft '{0}' auf der Auflistung vom Typ '{1}' nicht gefunden",
                [TemplateErrorCode.PropertyNotFoundOnType] = "Eigenschaft '{0}' nicht gefunden in '{1}' vom Typ '{2}'",
                [TemplateErrorCode.NotEnumerable] = "'{0}' ist keine Auflistung - der Wert ist vom Typ {1}",
                [TemplateErrorCode.RangeCountNotNumeric] = "'{0}' ist keine ganze Zahl - der Wert '{1}' kann nicht als ganze Zahl interpretiert werden",
                [TemplateErrorCode.RangeCountNotNumericOrEnumerable] = "'{0}' ist weder eine ganze Zahl noch eine Auflistung - der Wert ist vom Typ {1}",
                [TemplateErrorCode.NotADynamicTable] = "'{0}' ist nicht vom Typ {1}",
                [TemplateErrorCode.IndexingNotSupported] = "Ein Index kann nicht auf einen Ausdruck vom Typ '{0}' angewendet werden",

                // expressions
                [TemplateErrorCode.ScriptParseError] = "Fehler beim Parsen des Skripts {0}",
                [TemplateErrorCode.ExpressionParseError] = "Fehler beim Parsen des Ausdrucks {0}",
                [TemplateErrorCode.InvalidExpression] = "Ungültiger Ausdruck '{0}'",
                [TemplateErrorCode.ErrorInCondition] = "{0} in der Bedingung '{1}'",
                [TemplateErrorCode.ErrorInSwitchCase] = "{0} in Switch '{1}' Case '{2}'",

                // placeholder replacement
                [TemplateErrorCode.InvalidVariableSyntax] = "Ungültige Platzhalter-Syntax '{0}'",
                [TemplateErrorCode.PlaceholderNotReplaced] = "'{0}' konnte nicht ersetzt werden: {1} >> {0} << {2}",
                [TemplateErrorCode.ContentControlNotReplaced] = "Das Inhaltssteuerelement mit dem Tag '{0}' konnte nicht ersetzt werden: {1}",
                [TemplateErrorCode.ContentControlHasNoText] = "Der ermittelte Wert hat keinen Textlauf, der gefüllt werden könnte (leere Zellen-/Zeilen-Inhaltssteuerelemente werden nicht unterstützt)",

                // formatters
                [TemplateErrorCode.FormatPatternRequiresArgument] = "Der Formatierer FORMAT benötigt genau ein Argument, z.B. FORMAT(dd.MM.yyyy)",
                [TemplateErrorCode.FormatNotApplicable] = "Das Format {0} kann nicht auf {1} vom Typ {2} angewendet werden",
                [TemplateErrorCode.FormatterRequiresFormattable] = "Der Formatierer {0} kann nur auf IFormattable-Objekte angewendet werden - Eigenschaft {1}",
                [TemplateErrorCode.FormatterRequiresString] = "Der Formatierer {0} kann nur auf Zeichenketten angewendet werden - Eigenschaft {1}",
                [TemplateErrorCode.InvalidFormatterArguments] = "Argumente müssen die Form foo:bar haben",
                [TemplateErrorCode.SubTemplateNameRequired] = "Der Formatierer template benötigt einen Vorlagennamen",
                [TemplateErrorCode.SubTemplateInvalid] = "Die Vorlage ist null oder keine gültige OpenXML-Vorlage",
                [TemplateErrorCode.SubTemplateInvalidArgument] = "Ungültiges Argument '{0}' für den Formatierer template",
                [TemplateErrorCode.SubTemplateInvalidSelector] = "Ungültiger Selektor {0}",
                [TemplateErrorCode.SubTemplateUnsupportedElement] = "Die Vorlage muss ein Absatz, ein Textlauf, eine Tabellenzeile oder eine Tabellenzelle sein",
                [TemplateErrorCode.SubTemplateUnsupportedType] = "Die Vorlage muss ein string, OpenXmlElement, byte[] oder Stream sein",
                [TemplateErrorCode.HtmlNotInParagraph] = "Der HTML-Platzhalter befindet sich nicht in einem Absatz",
                [TemplateErrorCode.ImageMetadataUnreadable] = "Die Metadaten des Bildes konnten nicht gelesen werden",
                [TemplateErrorCode.InvalidImageFormatterArgument] = "Ungültiges Argument '{0}' für den Formatierer image",
                [TemplateErrorCode.MarkdownNestingDepthExceeded] = "Die maximale Verschachtelungstiefe von Markdown wurde überschritten",
                [TemplateErrorCode.MarkdownInvalidImage] = "Ungültige Bilddaten im Markdown-Link. {0}",

                // template syntax
                [TemplateErrorCode.TemplateSyntaxErrors] = "Syntaxfehler in der Vorlage:\n{0}",
                [TemplateErrorCode.SyntaxErrorLocation] = "{0}: {1} (bei '{2}')",
                [TemplateErrorCode.InvalidSyntax] = "Ungültige Syntax '{0}'",
                [TemplateErrorCode.UseGenericClosingTag] = "Ungültige Syntax '{0}'. Verwenden Sie stattdessen '{{{{/}}}}'.",
                [TemplateErrorCode.RangeLoopVariableNameRequired] = "Ungültige Syntax für die Zählschleife '{0}' - ein Variablenname ist erforderlich",
                [TemplateErrorCode.PatternMatchTimeout] = "Ungültige Syntax '{0}' - Zeitüberschreitung bei der Analyse",
                [TemplateErrorCode.UnknownKeyword] = "Unbekanntes Schlüsselwort '{0}'. Unterstützte Schlüsselwörter sind {1}",
                [TemplateErrorCode.CaseOutsideSwitch] = "'{0}' muss sich innerhalb eines '{{{{#switch: ...}}}}'-Blocks befinden",
                [TemplateErrorCode.ElseOutsideCondition] = "'{0}' (else) ist nur direkt innerhalb einer Bedingung '{{?{{...}}}}' erlaubt",
                [TemplateErrorCode.DuplicateElse] = "Die Bedingung '{0}' hat mehr als ein else '{1}'",
                [TemplateErrorCode.SeparatorOutsideLoop] = "Das Trennzeichen '{0}' ist nur direkt innerhalb einer Schleife über eine Auflistung '{{{{#Items}}}}' erlaubt",
                [TemplateErrorCode.DuplicateSeparator] = "Die Schleife '{0}' hat mehr als ein Trennzeichen '{1}'",
                [TemplateErrorCode.NoMatchingOpeningTag] = "'{0}' hat kein passendes öffnendes Tag",
                [TemplateErrorCode.BlockNotClosed] = "'{0}' ist nicht geschlossen",
                [TemplateErrorCode.CollectionNameRequired] = "'{0}' benötigt den Namen einer Auflistung, z.B. '{{{{#Items}}}}'",
                [TemplateErrorCode.SwitchExpressionRequired] = "'{0}' benötigt einen Ausdruck, z.B. '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.SwitchLooksLikeLoop] = "'{0}' ist eine Schleife über eine Auflistung namens '{1}'. Für ein switch / case verwenden Sie '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.ClosingTagMismatch] = "'{0}' passt nicht zu '{1}'",
                [TemplateErrorCode.ClosingTagExpectedGeneric] = "'{0}' schließt '{1}', erwartet wird '{{{{/}}}}'",
                [TemplateErrorCode.EmptyExpression] = "'{0}' hat einen leeren Ausdruck",
                [TemplateErrorCode.UnbalancedBracket] = "Unausgeglichene Klammer '{0}' im Ausdruck '{1}'",
                [TemplateErrorCode.UnterminatedStringLiteral] = "Nicht abgeschlossene Zeichenkette im Ausdruck '{0}'",
                [TemplateErrorCode.UnclosedBracket] = "Nicht geschlossene Klammer '{0}' im Ausdruck '{1}'",
                [TemplateErrorCode.UnexpectedClosingBraces] = "Unerwartetes '{0}' ohne passendes '{{{{'",
                [TemplateErrorCode.InvalidTag] = "Ungültiges Tag '{0}'",
                [TemplateErrorCode.TagNotTerminated] = "Das Tag '{0}' ist nicht abgeschlossen",
                [TemplateErrorCode.InvalidSwitchBlockSyntax] = "Ungültige Syntax für den Switch-Block: {0}",
                [TemplateErrorCode.InvalidCaseBlockSyntax] = "Ungültige Syntax für den Case-Block: {0}",
                [TemplateErrorCode.CaseNotNestedInSwitch] = "Der '{{#case}}'/'{{#default}}'-Block ('{0}') muss innerhalb eines '{{#switch}}'-Blocks liegen.",
                [TemplateErrorCode.InvalidInlineKeyword] = "Ungültiger Ausdruck {0}",
                [TemplateErrorCode.DynamicTableStructureInvalid] = "Ein dynamischer Tabellenblock muss genau eine Tabelle mit mindestens einer Kopf- und einer Datenzeile enthalten",
                [TemplateErrorCode.DynamicTableSingleChild] = "Ein dynamischer Tabellenblock muss genau einen untergeordneten Block enthalten"
            });
    }
}
