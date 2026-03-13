package pl.sk.cli.shell;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.springframework.shell.standard.FileValueProvider;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;
import org.springframework.shell.standard.ShellOption;
import pl.sk.cli.file.WordToMarkdown;

import java.io.IOException;

@ShellComponent
public class WordToMarkdownCommands {

    private final WordToMarkdown converter;

    public WordToMarkdownCommands(WordToMarkdown converter) {
        this.converter = converter;
    }

    @ShellMethod(key = "docx-to-md", value = "Convert DOCX file to Markdown")
    public String docxToMd(
            @ShellOption(help = "Input DOCX file path", valueProvider = FileValueProvider.class) String input,
            @ShellOption(help = "Output Markdown file path", defaultValue = ShellOption.NULL, valueProvider = FileValueProvider.class) String output
    ) {
        String outputPath = output;
        if (outputPath == null) {
            outputPath = input.replaceAll("\\.docx$", ".md");
        }

        try {
            String savedPath = converter.convertDocxToMarkdown(input, outputPath);
            return "Conversion completed. Markdown saved to: " + savedPath;
        } catch (Docx4JException | IOException e) {
            return "Conversion failed: " + e.getMessage();
        }
    }
}
