package pl.sk.cli.file;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.dml.CTBlip;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.picture.Pic;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.listnumbering.Emulator;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.P;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

@Service
public class WordToMarkdown {

    public String convertDocxToMarkdown(String inputFilePath, String outputFilePath) throws Docx4JException, IOException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputFilePath));
        String markdownContent = convertToMarkdown(wordMLPackage, Path.of(inputFilePath), Path.of(outputFilePath));
        saveMarkdownToFile(markdownContent, outputFilePath);
        return outputFilePath;
    }

    private String convertToMarkdown(WordprocessingMLPackage wordMLPackage, Path inputFilePath, Path outputFilePath) throws IOException {
        StringBuilder markdown = new StringBuilder();
        List<Object> documentContent = wordMLPackage.getMainDocumentPart().getContent();
        ConversionContext context = new ConversionContext(wordMLPackage, inputFilePath, outputFilePath);
        context.resetNumbering();

        for (int i = 0; i < documentContent.size(); i++) {
            Object content = unwrap(documentContent.get(i));
            if (content instanceof P) {
                P paragraph = (P) content;
                if (isCodeParagraph(paragraph)) {
                    i = handleCodeBlock(documentContent, i, markdown);
                } else {
                    handleParagraph(paragraph, markdown, context);
                }
            } else if (content instanceof Tbl) {
                Tbl table = (Tbl) content;
                handleTable(table, markdown, context);
            }
        }

        return normalizeMarkdown(markdown.toString());
    }

    private void handleParagraph(P paragraph, StringBuilder markdown, ConversionContext context) {
        String paragraphStyle = getParagraphStyle(paragraph);

        if (paragraphStyle != null && (paragraphStyle.startsWith("Heading") || paragraphStyle.startsWith("Nag"))) {
            int headingLevel = getHeadingLevel(paragraphStyle);
            markdown.append("#".repeat(headingLevel)).append(" ");
        }

        if (isListItem(paragraph)) {
            markdown.append(getListIndent(paragraph));
            String listPrefix = getListPrefix(paragraph, context);
            markdown.append(listPrefix).append(" ");
        }

        appendParagraphText(paragraph, markdown, context);
        markdown.append(isListItem(paragraph) ? "\n" : "\n\n");
    }

    private int handleCodeBlock(List<Object> documentContent, int startIndex, StringBuilder markdown) {
        P firstParagraph = (P) unwrap(documentContent.get(startIndex));
        String codeStyle = getParagraphStyle(firstParagraph);
        String codeLanguage = getCodeLanguage(codeStyle);

        markdown.append("```").append(codeLanguage).append("\n");

        int currentIndex = startIndex;
        while (currentIndex < documentContent.size()) {
            Object content = unwrap(documentContent.get(currentIndex));
            if (!(content instanceof P paragraph) || !codeStyle.equals(getParagraphStyle(paragraph))) {
                break;
            }

            appendPlainParagraphText(paragraph, markdown);
            markdown.append("\n");
            currentIndex++;
        }

        markdown.append("```").append("\n\n");
        return currentIndex - 1;
    }

    private void appendParagraphText(P paragraph, StringBuilder markdown, ConversionContext context) {
        List<Object> texts = paragraph.getContent();
        for (Object o : texts) {
            Object obj = unwrap(o);
            if (obj instanceof org.docx4j.wml.R) {
                org.docx4j.wml.R run = (org.docx4j.wml.R) obj;
                markdown.append(formatText(run, context));
            }
        }
    }

    private void appendPlainParagraphText(P paragraph, StringBuilder markdown) {
        List<Object> texts = paragraph.getContent();
        for (Object o : texts) {
            Object obj = unwrap(o);
            if (obj instanceof org.docx4j.wml.R run) {
                appendRunContent(run, markdown, false, null);
            }
        }
    }

    private boolean isListItem(P paragraph) {
        return paragraph.getPPr() != null && paragraph.getPPr().getNumPr() != null;
    }

    private String getListPrefix(P paragraph, ConversionContext context) {
        if (paragraph.getPPr() != null && paragraph.getPPr().getNumPr() != null) {
            String resolvedListPrefix = context.getListPrefix(paragraph);
            if (resolvedListPrefix != null) {
                return resolvedListPrefix;
            }
            return isNumberedList(paragraph) ? "1." : "-";
        }
        return "";
    }

    private String getListIndent(P paragraph) {
        if (!isListItem(paragraph) || paragraph.getPPr().getNumPr().getIlvl() == null) {
            return "";
        }

        int level = paragraph.getPPr().getNumPr().getIlvl().getVal().intValue();
        return "   ".repeat(Math.max(level, 0));
    }

    private boolean isNumberedList(P paragraph) {
        // Logic to determine if the list is numbered
        // Placeholder: refine based on actual Word numbering styles
        return paragraph.getPPr().getNumPr().getIlvl() == null
                || new BigInteger("0").equals(paragraph.getPPr().getNumPr().getIlvl().getVal());
    }

    private String getParagraphStyle(P paragraph) {
        if (paragraph.getPPr() != null && paragraph.getPPr().getPStyle() != null) {
            return paragraph.getPPr().getPStyle().getVal();
        }
        return null;
    }

    private boolean isCodeParagraph(P paragraph) {
        String paragraphStyle = getParagraphStyle(paragraph);
        return paragraphStyle != null
                && (paragraphStyle.startsWith("Code") || paragraphStyle.startsWith("Kod"));
    }

    private String getCodeLanguage(String paragraphStyle) {
        String suffix = paragraphStyle.startsWith("Code")
                ? paragraphStyle.substring("Code".length())
                : paragraphStyle.substring("Kod".length());
        return suffix.toLowerCase(Locale.ROOT);
    }

    private int getHeadingLevel(String paragraphStyle) {
        try {
            return Integer.parseInt(paragraphStyle.replaceAll("\\D", ""));
        } catch (NumberFormatException e) {
            return 1; // Default to H1 if parsing fails
        }
    }

    private void handleTable(Tbl table, StringBuilder markdown, ConversionContext context) {
        markdown.append("\n"); // Add a blank line before the table

        List<Object> rows = table.getContent();
        boolean isHeaderRow = true; // Flag to identify the header row

        for (Object r : rows) {
            Object rowObj = unwrap(r);
            if (rowObj instanceof Tr) {
                Tr row = (Tr) rowObj;
                handleTableRow(row, markdown, context);

                if (isHeaderRow) {
                    // Add separator row after the header row
                    int columnCount = row.getContent().size();
                    markdown.append('|');
                    for (int i = 0; i < columnCount; i++) {
                        markdown.append(" --- |");
                    }
                    markdown.append('\n');
                    isHeaderRow = false; // Only the first row is treated as the header
                }
            }
        }

        markdown.append("\n"); // Add a blank line after the table
    }

    private void handleTableRow(Tr row, StringBuilder markdown, ConversionContext context) {
        List<Object> cells = row.getContent();
        for (Object c : cells) {
            Object cellObj = unwrap(c);
            if (cellObj instanceof Tc) {
                Tc cell = (Tc) cellObj;
                String cellText = extractTextFromCell(cell, context);
                markdown.append("| ").append(cellText).append(" ");
            }
        }
        markdown.append("|\n");
    }

    private String extractTextFromCell(Tc cell, ConversionContext context) {
        StringBuilder cellText = new StringBuilder();
        List<Object> cellContent = cell.getContent();
        for (Object c : cellContent) {
            Object contentObj = unwrap(c);
            if (contentObj instanceof P) {
                P paragraph = (P) contentObj;
                appendParagraphText(paragraph, cellText, context);
                cellText.append(" ");
            }
        }
        return cellText.toString().trim();
    }

    private void saveMarkdownToFile(String markdownContent, String outputFilePath) throws IOException {
        Path outputPath = Path.of(outputFilePath);
        Files.writeString(
                outputPath,
                markdownContent,
                StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING
        );
    }

    private String normalizeMarkdown(String markdownContent) {
        String normalizedContent = markdownContent.replace("\r\n", "\n");
        String[] lines = normalizedContent.split("\n", -1);
        StringBuilder normalizedMarkdown = new StringBuilder();
        boolean inCodeBlock = false;
        boolean lastLineWasBlank = false;

        for (String line : lines) {
            if (line.startsWith("```")) {
                normalizedMarkdown.append(line).append("\n");
                inCodeBlock = !inCodeBlock;
                lastLineWasBlank = false;
                continue;
            }

            if (inCodeBlock) {
                normalizedMarkdown.append(line).append("\n");
                continue;
            }

            if (line.isBlank()) {
                if (normalizedMarkdown.length() == 0 || lastLineWasBlank) {
                    continue;
                }
                normalizedMarkdown.append("\n");
                lastLineWasBlank = true;
                continue;
            }

            normalizedMarkdown.append(line).append("\n");
            lastLineWasBlank = false;
        }

        return normalizedMarkdown.toString();
    }

    private String formatText(org.docx4j.wml.R run, ConversionContext context) {
        StringBuilder formattedText = new StringBuilder();
        RPr runProperties = run.getRPr();

        if (runProperties != null) {
            if (runProperties.getB() != null && runProperties.getB().isVal()) {
                formattedText.append("**"); // Bold
            }
            if (runProperties.getI() != null && runProperties.getI().isVal()) {
                formattedText.append("_"); // Italic
            }
            if (runProperties.getStrike() != null && runProperties.getStrike().isVal()) {
                formattedText.append("~~"); // Strikethrough
            }
            if (runProperties.getU() != null && runProperties.getU().getVal() != null) {
                formattedText.append("__"); // Underline
            }
        }

        appendRunContent(run, formattedText, true, context);

        if (runProperties != null) {
            if (runProperties.getU() != null && runProperties.getU().getVal() != null) {
                formattedText.append("__"); // Close underline
            }
            if (runProperties.getStrike() != null && runProperties.getStrike().isVal()) {
                formattedText.append("~~"); // Close strikethrough
            }
            if (runProperties.getI() != null && runProperties.getI().isVal()) {
                formattedText.append("_"); // Close italic
            }
            if (runProperties.getB() != null && runProperties.getB().isVal()) {
                formattedText.append("**"); // Close bold
            }
        }

        return formattedText.toString();
    }

    private void appendRunContent(org.docx4j.wml.R run, StringBuilder target, boolean includeImages, ConversionContext context) {
        run.getContent().forEach(c -> {
            Object content = unwrap(c);
            if (content instanceof Text text) {
                target.append(includeImages ? escapeMarkdownText(text.getValue()) : text.getValue());
            } else if (content instanceof Br) {
                target.append("\n");
            } else if (includeImages && content instanceof org.docx4j.wml.Drawing drawing) {
                target.append(extractImages(drawing, context));
            }
        });
    }

    private String escapeMarkdownText(String text) {
        return text
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }

    private String extractImages(org.docx4j.wml.Drawing drawing, ConversionContext context) {
        StringBuilder imagesMarkdown = new StringBuilder();
        for (Object anchorOrInline : drawing.getAnchorOrInline()) {
            Object drawingObject = unwrap(anchorOrInline);
            if (drawingObject instanceof Inline inline) {
                imagesMarkdown.append(exportImage(inline.getGraphic(), inline.getDocPr() != null ? inline.getDocPr().getDescr() : null, context));
            } else if (drawingObject instanceof Anchor anchor) {
                imagesMarkdown.append(exportImage(anchor.getGraphic(), anchor.getDocPr() != null ? anchor.getDocPr().getDescr() : null, context));
            }
        }
        return imagesMarkdown.toString();
    }

    private String exportImage(Graphic graphic, String altText, ConversionContext context) {
        if (graphic == null || graphic.getGraphicData() == null) {
            return "";
        }

        Pic pic = graphic.getGraphicData().getPic();
        if (pic == null || pic.getBlipFill() == null) {
            return "";
        }

        CTBlip blip = pic.getBlipFill().getBlip();
        if (blip == null) {
            return "";
        }

        String relationshipId = !blip.getEmbed().isEmpty() ? blip.getEmbed() : blip.getLink();
        if (relationshipId.isEmpty()) {
            return "";
        }

        try {
            String imagePath = context.exportImage(relationshipId);
            String resolvedAltText = (altText == null || altText.isBlank()) ? "image" : altText;
            return "![" + resolvedAltText + "](" + imagePath + ")";
        } catch (IOException e) {
            throw new IllegalStateException("Failed to export image for relationship " + relationshipId, e);
        }
    }

    private Object unwrap(final Object value) {
        if (value instanceof JAXBElement) {
            return ((JAXBElement<?>) value).getValue();
        }
        return value;
    }

    private static class ConversionContext {
        private final WordprocessingMLPackage wordMLPackage;
        private final Path resourcesDirectory;
        private final String resourcesDirectoryName;
        private final Map<String, String> imagePathsByRelationshipId = new HashMap<>();
        private final Set<String> usedFileNames = new HashSet<>();

        private ConversionContext(WordprocessingMLPackage wordMLPackage, Path inputFilePath, Path outputFilePath) {
            this.wordMLPackage = wordMLPackage;
            Path markdownDirectory = outputFilePath.toAbsolutePath().getParent();
            if (markdownDirectory == null) {
                markdownDirectory = Path.of(".").toAbsolutePath();
            }
            this.resourcesDirectoryName = createResourcesDirectoryName(inputFilePath);
            this.resourcesDirectory = markdownDirectory.resolve(resourcesDirectoryName);
        }

        private String exportImage(String relationshipId) throws IOException {
            if (imagePathsByRelationshipId.containsKey(relationshipId)) {
                return imagePathsByRelationshipId.get(relationshipId);
            }

            Part imagePart = resolveImagePart(relationshipId);
            if (!(imagePart instanceof BinaryPart binaryPart)) {
                throw new IOException("Relationship does not point to an embedded binary image: " + relationshipId);
            }

            Files.createDirectories(resourcesDirectory);
            String fileName = createUniqueFileName(imagePart);
            Path exportedImagePath = resourcesDirectory.resolve(fileName);
            Files.write(exportedImagePath, binaryPart.getBytes());

            String markdownPath = resourcesDirectoryName + "/" + fileName;
            imagePathsByRelationshipId.put(relationshipId, markdownPath);
            return markdownPath;
        }

        private String createResourcesDirectoryName(Path inputFilePath) {
            String fileName = inputFilePath.getFileName().toString();
            int extensionSeparator = fileName.lastIndexOf('.');
            String baseName = extensionSeparator > 0 ? fileName.substring(0, extensionSeparator) : fileName;
            return baseName.replaceAll("\\s+", "_") + "_resources";
        }

        private void resetNumbering() {
            NumberingDefinitionsPart numberingDefinitionsPart = wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart();
            if (numberingDefinitionsPart != null) {
                numberingDefinitionsPart.getEmulator(true);
            }
        }

        private String getListPrefix(P paragraph) {
            Emulator.ResultTriple result = Emulator.getNumber(wordMLPackage, paragraph.getPPr());
            if (result == null) {
                return null;
            }

            if (result.getBullet() != null && !result.getBullet().isBlank()) {
                return "-";
            }

            if (result.getNumString() == null || result.getNumString().isBlank()) {
                return null;
            }

            return result.getNumString();
        }

        private Part resolveImagePart(String relationshipId) {
            var relationshipsPart = wordMLPackage.getMainDocumentPart().getRelationshipsPart();
            var relationship = relationshipsPart.getRelationshipByID(relationshipId);
            return relationship == null ? null : relationshipsPart.getPart(relationship);
        }

        private String createUniqueFileName(Part imagePart) {
            String originalName = Path.of(imagePart.getPartName().getName()).getFileName().toString();
            String baseName = originalName;
            String extension = "";
            int extensionSeparator = originalName.lastIndexOf('.');
            if (extensionSeparator >= 0) {
                baseName = originalName.substring(0, extensionSeparator);
                extension = originalName.substring(extensionSeparator);
            }

            String candidate = originalName;
            int counter = 1;
            while (!usedFileNames.add(candidate)) {
                candidate = baseName + "-" + counter + extension;
                counter++;
            }
            return candidate;
        }
    }
}
