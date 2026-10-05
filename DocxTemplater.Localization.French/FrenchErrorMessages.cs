using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// French texts for the error messages of DocxTemplater.
    /// Register them once at startup: <c>TemplateErrorMessages.Default.AddFrench();</c>
    /// </summary>
    public static class FrenchErrorMessages
    {
        /// <summary>The neutral culture the texts are registered for; <c>fr-CH</c>, <c>fr-CA</c>, ... fall back to it.</summary>
        public static CultureInfo Culture { get; } = new("fr");

        /// <summary>
        /// Adds the French texts to <paramref name="messages"/>.
        /// </summary>
        /// <returns><paramref name="messages"/>, to chain calls.</returns>
        public static TemplateErrorMessages AddFrench(this TemplateErrorMessages messages)
        {
            return messages.AddLanguage(Culture, Formats);
        }

        public static IReadOnlyDictionary<TemplateErrorCode, string> Formats { get; } = new ReadOnlyDictionary<TemplateErrorCode, string>(
            new Dictionary<TemplateErrorCode, string>
            {
                // model binding
                [TemplateErrorCode.ModelPrefixAlreadyBound] = "Un modèle avec le préfixe '{0}' a déjà été lié - les préfixes de modèle ne distinguent pas les majuscules des minuscules",
                [TemplateErrorCode.ModelNotFound] = "Modèle {0} introuvable",
                [TemplateErrorCode.PropertyNotFound] = "Propriété {0} introuvable dans {1}",
                [TemplateErrorCode.PropertyNotFoundInParentScope] = "Propriété {0} introuvable dans la portée parente",
                [TemplateErrorCode.PropertyNotFoundOnCollection] = "Propriété '{0}' introuvable sur la collection de type '{1}'",
                [TemplateErrorCode.PropertyNotFoundOnType] = "Propriété '{0}' introuvable dans '{1}' de type '{2}'",
                [TemplateErrorCode.NotEnumerable] = "'{0}' n'est pas une collection - la valeur est de type {1}",
                [TemplateErrorCode.RangeCountNotNumeric] = "'{0}' n'est pas un nombre entier - la valeur '{1}' ne peut pas être interprétée comme un nombre entier",
                [TemplateErrorCode.RangeCountNotNumericOrEnumerable] = "'{0}' n'est ni un nombre entier ni une collection - la valeur est de type {1}",
                [TemplateErrorCode.NotADynamicTable] = "'{0}' n'est pas de type {1}",
                [TemplateErrorCode.IndexingNotSupported] = "Un index ne peut pas être appliqué à une expression de type '{0}'",

                // expressions
                [TemplateErrorCode.ScriptParseError] = "Erreur lors de l'analyse du script {0}",
                [TemplateErrorCode.ExpressionParseError] = "Erreur lors de l'analyse de l'expression {0}",
                [TemplateErrorCode.InvalidExpression] = "Expression invalide '{0}'",
                [TemplateErrorCode.ErrorInCondition] = "{0} dans la condition '{1}'",
                [TemplateErrorCode.ErrorInSwitchCase] = "{0} dans le switch '{1}' case '{2}'",

                // placeholder replacement
                [TemplateErrorCode.InvalidVariableSyntax] = "Syntaxe d'espace réservé invalide '{0}'",
                [TemplateErrorCode.PlaceholderNotReplaced] = "'{0}' n'a pas pu être remplacé : {1} >> {0} << {2}",
                [TemplateErrorCode.ContentControlNotReplaced] = "Le contrôle de contenu avec la balise '{0}' n'a pas pu être remplacé : {1}",
                [TemplateErrorCode.ContentControlHasNoText] = "La valeur résolue n'a aucun texte à remplir (les contrôles de contenu de cellule ou de ligne vides ne sont pas pris en charge)",

                // formatters
                [TemplateErrorCode.FormatPatternRequiresArgument] = "Le formateur FORMAT requiert exactement un argument, p. ex. FORMAT(dd.MM.yyyy)",
                [TemplateErrorCode.FormatNotApplicable] = "Le format {0} ne peut pas être appliqué à {1} de type {2}",
                [TemplateErrorCode.FormatterRequiresFormattable] = "Le formateur {0} ne peut être appliqué qu'à des objets IFormattable - propriété {1}",
                [TemplateErrorCode.FormatterRequiresString] = "Le formateur {0} ne peut être appliqué qu'à des chaînes de caractères - propriété {1}",
                [TemplateErrorCode.InvalidFormatterArguments] = "Les arguments doivent être de la forme foo:bar",
                [TemplateErrorCode.SubTemplateNameRequired] = "Le formateur template requiert un nom de modèle",
                [TemplateErrorCode.SubTemplateInvalid] = "Le modèle est null ou n'est pas un modèle OpenXML valide",
                [TemplateErrorCode.SubTemplateInvalidArgument] = "Argument '{0}' invalide pour le formateur template",
                [TemplateErrorCode.SubTemplateInvalidSelector] = "Sélecteur invalide {0}",
                [TemplateErrorCode.SubTemplateUnsupportedElement] = "Le modèle doit être un paragraphe, un run, une ligne de tableau ou une cellule de tableau",
                [TemplateErrorCode.SubTemplateUnsupportedType] = "Le modèle doit être un string, OpenXmlElement, byte[] ou Stream",
                [TemplateErrorCode.HtmlNotInParagraph] = "L'espace réservé HTML ne se trouve pas dans un paragraphe",
                [TemplateErrorCode.ImageMetadataUnreadable] = "Les métadonnées de l'image n'ont pas pu être lues",
                [TemplateErrorCode.InvalidImageFormatterArgument] = "Argument '{0}' invalide pour le formateur image",
                [TemplateErrorCode.MarkdownNestingDepthExceeded] = "La profondeur d'imbrication maximale du Markdown a été dépassée",
                [TemplateErrorCode.MarkdownInvalidImage] = "Données d'image invalides dans le lien Markdown. {0}",

                // template syntax
                [TemplateErrorCode.TemplateSyntaxErrors] = "Erreurs de syntaxe dans le modèle :\n{0}",
                [TemplateErrorCode.SyntaxErrorLocation] = "{0} : {1} (près de '{2}')",
                [TemplateErrorCode.InvalidSyntax] = "Syntaxe invalide '{0}'",
                [TemplateErrorCode.UseGenericClosingTag] = "Syntaxe invalide '{0}'. Utilisez '{{{{/}}}}' à la place.",
                [TemplateErrorCode.RangeLoopVariableNameRequired] = "Syntaxe de boucle invalide '{0}' - un nom de variable est requis",
                [TemplateErrorCode.PatternMatchTimeout] = "Syntaxe invalide '{0}' - délai d'analyse dépassé",
                [TemplateErrorCode.UnknownKeyword] = "Mot-clé inconnu '{0}'. Les mots-clés pris en charge sont {1}",
                [TemplateErrorCode.CaseOutsideSwitch] = "'{0}' doit se trouver dans un bloc '{{{{#switch: ...}}}}'",
                [TemplateErrorCode.ElseOutsideCondition] = "'{0}' (else) n'est autorisé que directement dans une condition '{{?{{...}}}}'",
                [TemplateErrorCode.DuplicateElse] = "La condition '{0}' a plus d'un else '{1}'",
                [TemplateErrorCode.SeparatorOutsideLoop] = "Le séparateur '{0}' n'est autorisé que directement dans une boucle sur une collection '{{{{#Items}}}}'",
                [TemplateErrorCode.DuplicateSeparator] = "La boucle '{0}' a plus d'un séparateur '{1}'",
                [TemplateErrorCode.NoMatchingOpeningTag] = "'{0}' n'a pas de balise ouvrante correspondante",
                [TemplateErrorCode.BlockNotClosed] = "'{0}' n'est pas fermé",
                [TemplateErrorCode.CollectionNameRequired] = "'{0}' requiert le nom d'une collection, p. ex. '{{{{#Items}}}}'",
                [TemplateErrorCode.SwitchExpressionRequired] = "'{0}' requiert une expression, p. ex. '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.SwitchLooksLikeLoop] = "'{0}' est une boucle sur une collection nommée '{1}'. Pour un switch / case, utilisez '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.ClosingTagMismatch] = "'{0}' ne correspond pas à '{1}'",
                [TemplateErrorCode.ClosingTagExpectedGeneric] = "'{0}' ferme '{1}', '{{{{/}}}}' est attendu",
                [TemplateErrorCode.EmptyExpression] = "'{0}' a une expression vide",
                [TemplateErrorCode.UnbalancedBracket] = "Parenthèse '{0}' non équilibrée dans l'expression '{1}'",
                [TemplateErrorCode.UnterminatedStringLiteral] = "Chaîne de caractères non terminée dans l'expression '{0}'",
                [TemplateErrorCode.UnclosedBracket] = "Parenthèse '{0}' non fermée dans l'expression '{1}'",
                [TemplateErrorCode.UnexpectedClosingBraces] = "'{0}' inattendu sans '{{{{' correspondant",
                [TemplateErrorCode.InvalidTag] = "Balise invalide '{0}'",
                [TemplateErrorCode.TagNotTerminated] = "La balise '{0}' n'est pas terminée",
                [TemplateErrorCode.InvalidSwitchBlockSyntax] = "Syntaxe invalide pour le bloc switch : {0}",
                [TemplateErrorCode.InvalidCaseBlockSyntax] = "Syntaxe invalide pour le bloc case : {0}",
                [TemplateErrorCode.CaseNotNestedInSwitch] = "Le bloc '{{#case}}'/'{{#default}}' ('{0}') doit être imbriqué dans un bloc '{{#switch}}'.",
                [TemplateErrorCode.InvalidInlineKeyword] = "Expression invalide {0}",
                [TemplateErrorCode.DynamicTableStructureInvalid] = "Un bloc de tableau dynamique doit contenir exactement un tableau avec au moins une ligne d'en-tête et une ligne de données",
                [TemplateErrorCode.DynamicTableSingleChild] = "Un bloc de tableau dynamique doit contenir exactement un bloc enfant"
            });
    }
}
