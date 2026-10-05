using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// Italian texts for the error messages of DocxTemplater.
    /// Register them once at startup: <c>TemplateErrorMessages.Default.AddItalian();</c>
    /// </summary>
    public static class ItalianErrorMessages
    {
        /// <summary>The neutral culture the texts are registered for; <c>it-CH</c>, <c>it-IT</c>, ... fall back to it.</summary>
        public static CultureInfo Culture { get; } = new("it");

        /// <summary>
        /// Adds the Italian texts to <paramref name="messages"/>.
        /// </summary>
        /// <returns><paramref name="messages"/>, to chain calls.</returns>
        public static TemplateErrorMessages AddItalian(this TemplateErrorMessages messages)
        {
            return messages.AddLanguage(Culture, Formats);
        }

        public static IReadOnlyDictionary<TemplateErrorCode, string> Formats { get; } = new ReadOnlyDictionary<TemplateErrorCode, string>(
            new Dictionary<TemplateErrorCode, string>
            {
                // model binding
                [TemplateErrorCode.ModelPrefixAlreadyBound] = "Un modello con il prefisso '{0}' è già stato associato - i prefissi dei modelli non distinguono tra maiuscole e minuscole",
                [TemplateErrorCode.ModelNotFound] = "Modello {0} non trovato",
                [TemplateErrorCode.PropertyNotFound] = "Proprietà {0} non trovata in {1}",
                [TemplateErrorCode.PropertyNotFoundInParentScope] = "Proprietà {0} non trovata nell'ambito superiore",
                [TemplateErrorCode.PropertyNotFoundOnCollection] = "Proprietà '{0}' non trovata nella raccolta di tipo '{1}'",
                [TemplateErrorCode.PropertyNotFoundOnType] = "Proprietà '{0}' non trovata in '{1}' di tipo '{2}'",
                [TemplateErrorCode.NotEnumerable] = "'{0}' non è una raccolta - il valore è di tipo {1}",
                [TemplateErrorCode.RangeCountNotNumeric] = "'{0}' non è un numero intero - il valore '{1}' non può essere interpretato come numero intero",
                [TemplateErrorCode.RangeCountNotNumericOrEnumerable] = "'{0}' non è né un numero intero né una raccolta - il valore è di tipo {1}",
                [TemplateErrorCode.NotADynamicTable] = "'{0}' non è di tipo {1}",
                [TemplateErrorCode.IndexingNotSupported] = "Non è possibile applicare un indice a un'espressione di tipo '{0}'",

                // expressions
                [TemplateErrorCode.ScriptParseError] = "Errore durante l'analisi dello script {0}",
                [TemplateErrorCode.ExpressionParseError] = "Errore durante l'analisi dell'espressione {0}",
                [TemplateErrorCode.InvalidExpression] = "Espressione non valida '{0}'",
                [TemplateErrorCode.ErrorInCondition] = "{0} nella condizione '{1}'",
                [TemplateErrorCode.ErrorInSwitchCase] = "{0} nello switch '{1}' case '{2}'",

                // placeholder replacement
                [TemplateErrorCode.InvalidVariableSyntax] = "Sintassi del segnaposto non valida '{0}'",
                [TemplateErrorCode.PlaceholderNotReplaced] = "'{0}' non può essere sostituito: {1} >> {0} << {2}",
                [TemplateErrorCode.ContentControlNotReplaced] = "Il controllo contenuto con il tag '{0}' non può essere sostituito: {1}",
                [TemplateErrorCode.ContentControlHasNoText] = "Il valore risolto non ha alcun testo da riempire (i controlli contenuto di cella o riga vuoti non sono supportati)",

                // formatters
                [TemplateErrorCode.FormatPatternRequiresArgument] = "Il formattatore FORMAT richiede esattamente un argomento, ad es. FORMAT(dd.MM.yyyy)",
                [TemplateErrorCode.FormatNotApplicable] = "Il formato {0} non può essere applicato a {1} di tipo {2}",
                [TemplateErrorCode.FormatterRequiresFormattable] = "Il formattatore {0} può essere applicato solo a oggetti IFormattable - proprietà {1}",
                [TemplateErrorCode.FormatterRequiresString] = "Il formattatore {0} può essere applicato solo a stringhe - proprietà {1}",
                [TemplateErrorCode.InvalidFormatterArguments] = "Gli argomenti devono essere nella forma foo:bar",
                [TemplateErrorCode.SubTemplateNameRequired] = "Il formattatore template richiede il nome di un modello",
                [TemplateErrorCode.SubTemplateInvalid] = "Il modello è null o non è un modello OpenXML valido",
                [TemplateErrorCode.SubTemplateInvalidArgument] = "Argomento '{0}' non valido per il formattatore template",
                [TemplateErrorCode.SubTemplateInvalidSelector] = "Selettore non valido {0}",
                [TemplateErrorCode.SubTemplateUnsupportedElement] = "Il modello deve essere un paragrafo, un run, una riga di tabella o una cella di tabella",
                [TemplateErrorCode.SubTemplateUnsupportedType] = "Il modello deve essere uno string, OpenXmlElement, byte[] o Stream",
                [TemplateErrorCode.HtmlNotInParagraph] = "Il segnaposto HTML non si trova in un paragrafo",
                [TemplateErrorCode.ImageMetadataUnreadable] = "Impossibile leggere i metadati dell'immagine",
                [TemplateErrorCode.InvalidImageFormatterArgument] = "Argomento '{0}' non valido per il formattatore image",
                [TemplateErrorCode.MarkdownNestingDepthExceeded] = "È stata superata la profondità massima di annidamento del Markdown",
                [TemplateErrorCode.MarkdownInvalidImage] = "Dati immagine non validi nel link Markdown. {0}",

                // template syntax
                [TemplateErrorCode.TemplateSyntaxErrors] = "Errori di sintassi nel modello:\n{0}",
                [TemplateErrorCode.SyntaxErrorLocation] = "{0}: {1} (vicino a '{2}')",
                [TemplateErrorCode.InvalidSyntax] = "Sintassi non valida '{0}'",
                [TemplateErrorCode.UseGenericClosingTag] = "Sintassi non valida '{0}'. Utilizzare invece '{{{{/}}}}'.",
                [TemplateErrorCode.RangeLoopVariableNameRequired] = "Sintassi del ciclo non valida '{0}' - è richiesto il nome di una variabile",
                [TemplateErrorCode.PatternMatchTimeout] = "Sintassi non valida '{0}' - timeout durante l'analisi",
                [TemplateErrorCode.UnknownKeyword] = "Parola chiave sconosciuta '{0}'. Le parole chiave supportate sono {1}",
                [TemplateErrorCode.CaseOutsideSwitch] = "'{0}' deve trovarsi all'interno di un blocco '{{{{#switch: ...}}}}'",
                [TemplateErrorCode.ElseOutsideCondition] = "'{0}' (else) è consentito solo direttamente all'interno di una condizione '{{?{{...}}}}'",
                [TemplateErrorCode.DuplicateElse] = "La condizione '{0}' ha più di un else '{1}'",
                [TemplateErrorCode.SeparatorOutsideLoop] = "Il separatore '{0}' è consentito solo direttamente all'interno di un ciclo su una raccolta '{{{{#Items}}}}'",
                [TemplateErrorCode.DuplicateSeparator] = "Il ciclo '{0}' ha più di un separatore '{1}'",
                [TemplateErrorCode.NoMatchingOpeningTag] = "'{0}' non ha un tag di apertura corrispondente",
                [TemplateErrorCode.BlockNotClosed] = "'{0}' non è chiuso",
                [TemplateErrorCode.CollectionNameRequired] = "'{0}' richiede il nome di una raccolta, ad es. '{{{{#Items}}}}'",
                [TemplateErrorCode.SwitchExpressionRequired] = "'{0}' richiede un'espressione, ad es. '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.SwitchLooksLikeLoop] = "'{0}' è un ciclo su una raccolta denominata '{1}'. Per uno switch / case utilizzare '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.ClosingTagMismatch] = "'{0}' non corrisponde a '{1}'",
                [TemplateErrorCode.ClosingTagExpectedGeneric] = "'{0}' chiude '{1}', è atteso '{{{{/}}}}'",
                [TemplateErrorCode.EmptyExpression] = "'{0}' ha un'espressione vuota",
                [TemplateErrorCode.UnbalancedBracket] = "Parentesi '{0}' non bilanciata nell'espressione '{1}'",
                [TemplateErrorCode.UnterminatedStringLiteral] = "Stringa non terminata nell'espressione '{0}'",
                [TemplateErrorCode.UnclosedBracket] = "Parentesi '{0}' non chiusa nell'espressione '{1}'",
                [TemplateErrorCode.UnexpectedClosingBraces] = "'{0}' inatteso senza '{{{{' corrispondente",
                [TemplateErrorCode.InvalidTag] = "Tag non valido '{0}'",
                [TemplateErrorCode.TagNotTerminated] = "Il tag '{0}' non è terminato",
                [TemplateErrorCode.InvalidSwitchBlockSyntax] = "Sintassi non valida per il blocco switch: {0}",
                [TemplateErrorCode.InvalidCaseBlockSyntax] = "Sintassi non valida per il blocco case: {0}",
                [TemplateErrorCode.CaseNotNestedInSwitch] = "Il blocco '{{#case}}'/'{{#default}}' ('{0}') deve essere annidato in un blocco '{{#switch}}'.",
                [TemplateErrorCode.InvalidInlineKeyword] = "Espressione non valida {0}",
                [TemplateErrorCode.DynamicTableStructureInvalid] = "Un blocco di tabella dinamica deve contenere esattamente una tabella con almeno una riga di intestazione e una riga di dati",
                [TemplateErrorCode.DynamicTableSingleChild] = "Un blocco di tabella dinamica deve contenere esattamente un blocco figlio"
            });
    }
}
