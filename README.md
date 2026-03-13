# Word to Markdown

`word-to-markdown` is a small command-line tool that converts `.docx` files to Markdown.

The application is built with:

- Java 17
- Spring Boot
- Spring Shell
- docx4j

## What It Does

The converter reads the main content of a Microsoft Word `.docx` document and generates a Markdown file.

Current conversion support includes:

- regular paragraphs
- headings based on Word styles such as `Heading` and `Nag`
- bulleted lists
- nested bulleted and numbered lists based on Word list levels
- numbered lists with generated source numbering
- fenced code blocks based on paragraph styles starting with `Code` or `Kod`
- code block language detection from style suffixes, for example `CodeJava` -> `java`
- grouped multi-paragraph code blocks when adjacent paragraphs use the same code style
- inline formatting for bold, italic, underline, and strikethrough
- Markdown tables
- embedded images exported to a `resources` directory and linked from Markdown
- output normalization that removes leading blank lines and collapses multiple blank lines outside code blocks

## Requirements

- Java 17 or newer
- Maven 3.9+ recommended

## Build

```bash
mvn clean compile
```

## Run

You can start the shell with Maven:

```bash
mvn spring-boot:run
```

Or pass application arguments directly:

```bash
mvn spring-boot:run "-Dspring-boot.run.arguments=--help"
```

After startup, use the Spring Shell command:

```text
docx-to-md --input <input.docx> --output <output.md>
```

Example:

```text
docx-to-md --input my-document.docx --output my-document.md
```

If `--output` is omitted, the tool saves the result next to the input file using the same base name and the `.md` extension.

## Output Structure

The converter creates:

- the target Markdown file
- a `resources` directory next to the Markdown file when the document contains embedded images

Example:

```text
docs/
  input.docx
  output.md
  resources/
    image1.png
    image2.jpeg
```

In the generated Markdown, images are referenced like this:

```md
![image](resources/image1.png)
```

## Code Block Styles

Paragraph styles beginning with `Code` or `Kod` are exported as fenced Markdown code blocks.

Examples:

- `Code` / `Kod` -> plain fenced block
- `CodeJava` / `KodSQL` -> `java`

If multiple consecutive paragraphs use the same code style, they are merged into a single fenced block.

Example:

````md
```java
public static void main(String[] args) {
}
```
````

## Supported Command

The application currently exposes one shell command:

### `docx-to-md`

Convert a Word document to Markdown.

Arguments:

- `--input` - path to the source `.docx` file
- `--output` - optional path to the generated `.md` file

## Behavior Notes

- Paragraphs are separated by a single blank line in the generated Markdown.
- List items are written line by line without extra blank lines between items.
- Nested lists are indented using Markdown-compatible spacing derived from Word list levels.
- Code blocks preserve line-by-line paragraph content.
- Blank lines at the beginning of the generated file are removed.
- Multiple consecutive blank lines outside fenced code blocks are collapsed to one blank line.
- Image file names are deduplicated when necessary.

## Known Limitations

The current implementation is intentionally simple and does not attempt full Word-to-Markdown fidelity.

Known limitations include:

- only the main document body is processed
- complex list formatting may not exactly match every Word numbering scheme
- inline Word formatting is limited to bold, italic, underline, and strikethrough
- tables are exported as basic Markdown tables
- image links are resolved from embedded document relationships in the main document part
- advanced Word content such as footnotes, comments, headers, footers, text boxes, SmartArt, charts, and tracked changes is not specifically handled

## Development

Compile the project:

```bash
mvn compile
```

Run the application:

```bash
mvn spring-boot:run
```

Package the application:

```bash
mvn package
```

## Project Structure

Key files:

- `src/main/java/pl/sk/cli/file/WordToMarkdown.java` - conversion logic
- `src/main/java/pl/sk/cli/shell/WordToMarkdownCommands.java` - shell command definition
- `src/main/java/pl/sk/cli/Application.java` - Spring Boot entry point
