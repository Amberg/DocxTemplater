# DocxTemplater

_DocxTemplater is a library to generate docx documents from a docx template. 
The template can be **bound to multiple datasources** and be edited by non-programmers.
It supports placeholder **replacement** and **loops** and **images**_

**Features:**
* Variable Replacement
* Loops - Bind to collections
* Conditional Blocks
* HTML Snippets - Replace placeholder with HTML Content
* Images - Replace placeholder with Image data


## Quickstart

Create a ocx template with placeholder syntax

```
This Text: {{ds.Title}} - will be replaced
```
To open the template, add a model and store the result to a file
```c#
ver template = DocxTemplate.Open("template.docx")
template.AddModel("ds", new {Title = "Some Text"})
template.ProcessToFile("generated.docx")
```
The generated word document then contains

```
This Text: Some Text - will be replaced
```


## Placeholder Syntax

A placholder can consist of three parts: {{**property**}:**formatter**(**arguments**)}

- **property**:   the path to the property in the datasource objects.
- **formatter**:  formatter applied to convert the model value to openxml _(ae. toupper, tolower img format etc)_ 
- **arguments**: formatter arguments - some formatter have arguments

The syntax is case insensitive

**Quick Reference:** (Expamples)

|      Syntax      |               Desciption |
|----------------|--------------------------|
| {{SomeVar}}  | Simple Variable replacement
| {{someVar > 5}}...{{else}}...{{/}}  | Conditional blocks
| {{/Items}}...{\{Items.Name}} ... {{/Items}}  | Text block bound to collection items
| {{SomeString:ToUpper()}}  | Variable with formatter to upper
| {{SomeDate:Format("MM/dd/yyyy")}}  | Date variable with formatting
| {{SomeDate:F("MM/dd/yyyy")}}  | Date variable with formatting - short syntax
| {{SomeBytes:img()}}  | Image Formatter for image data
| {{SomeHtmlString:html()}}  | Inserts html string into word document
### Loops

To repeat document content for each item in a collection the loop syntax can be used:
**{{#_\<collection\>_}}** .. content .. **{{_<collection\>_}}**
All document content between the start and end tag is rendered for each element in the collection. 

```
{{#Items}} This text {{Items.Name}} is rendered for each element in the items collection {{/items}}
```

This can be used, for example, to bind a collection to a table. In this case, the start and end tag has to be placed in the row of the table

|      Name      | Position |
|----------------|----------|
| **{{#Items}}** {{Items.Name}} | {{Items.Position}} **{{/Items}}*|

This template bound to a model:
```c#
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

will render a table row for each item in the collection

|      Name      | Position |
|----------------|----------|
| John | Developer|
| Alice | CEO|

### Conditional Blocks

Show or hide a given section depending on a condition:
**{{\<condition>}}** .. content .. **{{/}}**
All document content between the start and end tag is rendered only if the condition is met

```
{{Item.Value >= 0}}Only visible if value is >= 0 {{/}}
{{Item.Value < 0}}Otherwise this text is shown{{/}}
```

## Formatters

If no formatter is specified, the model value is converted into a text with "ToString".

This is not sufficient for all data types. That is why there are formatters that convert text or binary data into the desired representation

The formatter name is always case insensitive

### String Formatters

- ToUpper, ToLower

### FormatPatterns

Any type that implements ```IFormattable``` can be formatted with the net formatter strings

**See:**
[Standard date and time format strings](https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-date-and-time-format-strings)
[Standard numeric format strings](https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings)
.. and many more

**Examples:**
{{SomeDate}:format(d)}  ----> "6/15/2009"  (en-US)
{{SomeDouble}:format(f2)}  ----> "1234.42"  (en-US)

### Image Formatter

---

**_NOTE:_** for the Image formatter the nuget package *DocxTemplater.Images* is required text

---

Because the image formatter is not standard, it must be added
```c#
var docTemplate = new DocxTemplate(fileStream);
docTemplate.RegisterFormatter(new ImageFormatter());
```

The image formatter replaces a placeholder with an image stored in a byte array 

The placeholder can be placed in a TextBox so that the end user can easily adjust the image size in the template. The size of the image is then adapted to the size of the TextBox.

The stretching behavior can be configured

|      Arg      | Example | Description
|----------------|----------|---
| KEEPRATIO| {{imgData}:img(keepratio)} | Scales the image to fit the container - keeps aspect ratio
| STRETCHW | {imgData}:img(STRETCHW)}| Scales the image to fit the width of the container
| STRETCHH | {imgData}:img(STRETCHH)}| Scales the image to fit the height of the container
