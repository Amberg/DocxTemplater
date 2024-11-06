To update the README, you can edit the file directly in the repository. Here is a revised version of the README content you can use:

---

# DocxTemplater

_DocxTemplater is a library to generate docx documents from a docx template. The template can be **bound to multiple datasources** and be edited by non-programmers. It supports placeholder **replacement**, **loops**, and **images**._

[![CI-Build](https://github.com/Amberg/DocxTemplater/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/Amberg/DocxTemplater/actions/workflows/ci.yml)

## Features
- Variable Replacement
- Collections - Bind to collections
- Conditional Blocks
- Dynamic Tables - Columns are defined by the datasource
- Markdown Support - Converts Markdown to OpenXML
- HTML Snippets - Replace placeholder with HTML Content
- Images - Replace placeholder with Image data

## Quickstart

Create a docx template with placeholder syntax:
```
This Text: {{ds.Title}} - will be replaced
```

Open the template, add a model, and store the result to a file:
```csharp
var template = DocxTemplate.Open("template.docx");
// To open the file from a stream use the constructor directly 
// var template = new DocxTemplate(stream);
template.BindModel("ds", new { Title = "Some Text" });
template.Save("generated.docx");
```

The generated word document will contain:
```
This Text: Some Text - will be replaced
```

### Install DocxTemplater via NuGet

To include DocxTemplater in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/DocxTemplater).

Run the following command in the Package Manager Console:
```
PM> Install-Package DocxTemplater
```

For Image support:
```
PM> Install-Package DocxTemplater.Images
```

## Placeholder Syntax

A placeholder can consist of three parts: `{{**property**:**formatter**(**arguments**)}}`

- **property**: The path to the property in the datasource objects.
- **formatter**: Formatter applied to convert the model value to OpenXML (e.g., `toupper`, `tolower`, `img` format).
- **arguments**: Formatter arguments - some formatters have arguments.

The syntax is case insensitive.

### Quick Reference Examples

| Syntax                              | Description                                                    |
|-------------------------------------|----------------------------------------------------------------|
| `{{SomeVar}}`                       | Simple Variable replacement                                    |
| `{?{someVar > 5}}...{{:}}...{{/}}`  | Conditional blocks                                             |
| `{{#Items}}...{{Items.Name}} ... {{/Items}}` | Text block bound to collection of complex items          |
| `{{#Items}}...{{.Name}} ... {{/Items}}` | Same as above with dot notation - implicit iterator          |
| `{{#Items}}...{{.}:toUpper} ... {{/Items}}` | A list of string all upper case - dot notation              |
| `{{#Items}}{{.}}{{:s:}},{{/Items}}` | A list of strings comma separated - dot notation              |
| `{{SomeString}:ToUpper()}`          | Variable with formatter to upper                               |
| `{{SomeDate}:Format('MM/dd/yyyy')}` | Date variable with formatting                                  |
| `{{SomeDate}:F('MM/dd/yyyy')}`      | Date variable with formatting - short syntax                   |
| `{{SomeBytes}:img()}`               | Image Formatter for image data                                 |
| `{{SomeHtmlString}:html()}`         | Inserts HTML string into the word document                     |

### Collections

To repeat document content for each item in a collection, use the loop syntax:
`**{{#_\<collection\>_}}** ... content ... **{{/_\<collection\>_}}**`

All document content between the start and end tag is rendered for each element in the collection:
```
{{#Items}} This text {{Items.Name}} is rendered for each element in the items collection {{/Items}}
```

This can be used, for example, to bind a collection to a table. In this case, the start and end tag have to be placed in the row of the table:
| Name         | Position  |
|--------------|-----------|
| **{{#Items}}** {{Items.Name}} | {{Items.Position}} **{{/Items}}** |

This template bound to a model:
```csharp
var template = DocxTemplate.Open("template.docx");
var model = new
{
    Items = new[]
    {
        new { Name = "John", Position = "Developer" },
        new { Name = "Alice", Position = "CEO" }
    }
};
template.BindModel("ds", model);
template.Save("generated.docx");
```

Will render a table row for each item in the collection:
| Name  | Position  |
|-------|-----------|
| John  | Developer |
| Alice | CEO       |

#### Separator

To render a separator between the items in the collection, use the separator syntax:
```
{{#Items}} This text {{.Name}} is rendered for each element in the items collection {{:s:}} This is rendered between each element {{/Items}}
```

### Conditional Blocks

Show or hide a given section depending on a condition:
`**{?{\<condition\>}}** ... content ... **{{/}}**`

All document content between the start and end tag is rendered only if the condition is met:
```
{?{Item.Value >= 0}}Only visible if value is >= 0
{{:}}Otherwise this text is shown{{/}}
```

## Formatters

If no formatter is specified, the model value is converted into a text with `ToString`.

This is not sufficient for all data types. That is why there are formatters that convert text or binary data into the desired representation.

The formatter name is always case insensitive.

### String Formatters

- `ToUpper`
- `ToLower`

### FormatPatterns

Any type that implements `IFormattable` can be formatted with the standard format strings for this type.

See:
- [Standard date and time format strings](https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-date-and-time-format-strings)
- [Standard numeric format strings](https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings)

Examples:
```
{{SomeDate}:format(d)}  ----> "6/15/2009"  (en-US)
{{SomeDouble}:format(f2)}  ----> "1234.42"  (en-US)
```

### Image Formatter

**_NOTE:_** For the Image formatter, the NuGet package `DocxTemplater.Images` is required.

Because the image formatter is not standard, it must be added:
```csharp
var docTemplate = new DocxTemplate(fileStream);
docTemplate.RegisterFormatter(new ImageFormatter());
```

The image formatter replaces a placeholder with an image stored in a byte array.

The placeholder can be placed in a TextBox so that the end user can easily adjust the image size in the template. The size of the image is then adapted to the size of the TextBox.

The stretching behavior can be configured:
| Arg        | Example                         | Description                                      |
|------------|---------------------------------|--------------------------------------------------|
| `KEEPRATIO`| `{{imgData}:img(keepratio)}`    | Scales the image to fit the container - keeps aspect ratio |
| `STRETCHW` | `{imgData}:img(STRETCHW)}`      | Scales the image to fit the width of the container |
| `STRETCHH` | `{imgData}:img(STRETCHH)}`      | Scales the image to fit the height of the container |

If the image is used without any container the image scaling can be with 'w' or 'h' argument. Use 'r' to rotate the image.
Is only one of the arguments 'w' or 'h' used the image is scaled to the given width or height and the aspect ratio is kept.
The size of the image can be specified in different units (cm,mm,in,px)

|      Arg      | Example | Description
|----------------|----------|---
| w | `{{imgData}:img(w:100mm)}` | Scales the image to a width of 100 millimeters
| h | `{{imgData}:img(h:100in)}` | Scales the image to a height of 100 inches
| r | `{{imgData}:img(r:90)}`    | Rotates the image by 90 degrees

### Error Handling

If a placeholder is not found in the model, an exception is thrown. This can be configured with the `ProcessSettings`:
```csharp
var docTemplate = new DocxTemplate(memStream);
docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent;
var result = docTemplate.Process();
```

### Culture

The culture used to format the model values can be configured with the `ProcessSettings`:
```csharp
var docTemplate = new DocxTemplate(memStream, new ProcessSettings()
{
    Culture = new CultureInfo("en-us")
});
var result = docTemplate.Process();
```