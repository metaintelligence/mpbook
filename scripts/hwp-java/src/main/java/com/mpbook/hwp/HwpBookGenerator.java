package com.mpbook.hwp;

import com.google.gson.Gson;
import com.google.gson.annotations.SerializedName;
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;

import java.io.IOException;
import java.io.Reader;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class HwpBookGenerator {
    private static final Gson GSON = new Gson();
    private static final Pattern INLINE_FORMULA_PATTERN = Pattern.compile("\\\\\\((.+?)\\\\\\)");

    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);

        Path inputJson = Paths.get(parsed.inputJsonPath).toAbsolutePath().normalize();
        Path outputHwp = Paths.get(parsed.outputHwpPath).toAbsolutePath().normalize();
        Path visualDir = Paths.get(parsed.visualDirPath).toAbsolutePath().normalize();

        QuestionPayload payload = loadPayload(inputJson);
        HWPFile hwpFile = buildBook(payload, visualDir);

        Files.createDirectories(outputHwp.getParent());
        HWPWriter.toFile(hwpFile, outputHwp.toString());

        System.out.println("Generated HWP (safe mode): " + outputHwp);
    }

    private static QuestionPayload loadPayload(Path inputJson) throws IOException {
        try (Reader reader = Files.newBufferedReader(inputJson, StandardCharsets.UTF_8)) {
            QuestionPayload payload = GSON.fromJson(reader, QuestionPayload.class);
            if (payload == null) {
                throw new IllegalArgumentException("JSON payload is null: " + inputJson);
            }
            if (payload.questions == null) {
                payload.questions = Collections.emptyList();
            }
            return payload;
        }
    }

    private static HWPFile buildBook(QuestionPayload payload, Path visualDir) throws UnsupportedEncodingException {
        HWPFile hwpFile = BlankFileMaker.make();
        Section section = hwpFile.getBodyText().getSectionList().get(0);

        addParagraph(section, "Math Problem Book (HWP Safe Mode)");
        addParagraph(section, "Generated from structured_questions.json");
        addParagraph(section, "");

        int index = 0;
        for (Question q : safeQuestions(payload.questions)) {
            index++;
            String qid = valueOrFallback(q.id, String.format("Q%02d", index));
            String qtype = valueOrFallback(q.type, "unknown");
            String title = valueOrFallback(q.title, "");

            if (title.isEmpty()) {
                addParagraph(section, qid + " [" + qtype + "]");
            } else {
                addParagraph(section, qid + " [" + qtype + "] " + title);
            }

            for (String line : safeLines(q.body)) {
                addParagraph(section, normalizeRichText(line));
            }

            for (String visualHint : resolveVisualHints(visualDir, qid)) {
                addParagraph(section, visualHint);
            }

            List<String> choices = safeLines(q.choices);
            if (!choices.isEmpty()) {
                addParagraph(section, "Choices:");
                int cIdx = 1;
                for (String c : choices) {
                    addParagraph(section, cIdx + ". " + normalizeRichText(c));
                    cIdx++;
                }
            }

            addParagraph(section, "");
            addParagraph(section, "");
        }

        return hwpFile;
    }

    private static List<Question> safeQuestions(List<Question> questions) {
        if (questions == null) {
            return Collections.emptyList();
        }
        return questions;
    }

    private static List<String> safeLines(List<String> lines) {
        if (lines == null) {
            return Collections.emptyList();
        }
        return lines;
    }

    private static String valueOrFallback(String value, String fallback) {
        if (value == null || value.isBlank()) {
            return fallback;
        }
        return value;
    }

    private static String normalizeRichText(String input) {
        if (input == null) {
            return "";
        }
        String normalized = input.replace("\r\n", " ").replace('\r', ' ').replace('\n', ' ').replace('\t', ' ');

        Matcher matcher = INLINE_FORMULA_PATTERN.matcher(normalized);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            String inner = matcher.group(1) == null ? "" : matcher.group(1).trim();
            matcher.appendReplacement(sb, Matcher.quoteReplacement(inner));
        }
        matcher.appendTail(sb);

        return sb.toString().replaceAll("\\s+", " ").trim();
    }

    private static void addParagraph(Section section, String text) throws UnsupportedEncodingException {
        Paragraph paragraph = section.addNewParagraph();

        paragraph.getHeader().setParaShapeId(3);
        paragraph.getHeader().setStyleId((short) 0);
        paragraph.getHeader().getDivideSort().setValue((short) 0);
        paragraph.getHeader().setInstanceID(0);
        paragraph.getHeader().setIsMergedByTrack(0);

        paragraph.createText();
        paragraph.getText().addString(text == null ? "" : text);

        paragraph.createCharShape();
        paragraph.getCharShape().addParaCharShape(0, 0);
    }

    private static List<String> resolveVisualHints(Path visualDir, String qid) {
        if (qid == null || qid.isBlank() || visualDir == null || !Files.isDirectory(visualDir)) {
            return Collections.emptyList();
        }

        List<Path> matches = new ArrayList<>();
        String[] patterns = {
                qid + "_*.png",
                qid + "_*.jpg",
                qid + "_*.jpeg"
        };

        for (String pattern : patterns) {
            try (DirectoryStream<Path> ds = Files.newDirectoryStream(visualDir, pattern)) {
                for (Path p : ds) {
                    matches.add(p.toAbsolutePath().normalize());
                }
            } catch (IOException ignored) {
            }
        }

        if (matches.isEmpty()) {
            return Collections.emptyList();
        }

        matches.sort(Comparator.comparing(Path::toString));
        List<String> hints = new ArrayList<>();
        for (Path p : matches) {
            hints.add("[visual] " + p.getFileName());
        }
        return hints;
    }

    private static final class Arguments {
        private final String inputJsonPath;
        private final String outputHwpPath;
        private final String visualDirPath;

        private Arguments(String inputJsonPath, String outputHwpPath, String visualDirPath) {
            this.inputJsonPath = inputJsonPath;
            this.outputHwpPath = outputHwpPath;
            this.visualDirPath = visualDirPath;
        }

        private static Arguments parse(String[] args) {
            String input = "output/structured_questions.json";
            String output = "output/math_problem_book_oss.hwp";
            String visualDir = "output/derived_visuals";

            for (int i = 0; i < args.length; i++) {
                String arg = args[i];
                if ("--input".equals(arg) && i + 1 < args.length) {
                    input = args[++i];
                } else if ("--output".equals(arg) && i + 1 < args.length) {
                    output = args[++i];
                } else if ("--visual-dir".equals(arg) && i + 1 < args.length) {
                    visualDir = args[++i];
                }
            }
            return new Arguments(input, output, visualDir);
        }
    }

    private static final class QuestionPayload {
        String project;
        @SerializedName("created_at")
        String createdAt;
        @SerializedName("source_folder")
        String sourceFolder;
        List<Question> questions;
    }

    private static final class Question {
        String id;
        @SerializedName("source_image")
        String sourceImage;
        String type;
        String title;
        List<String> body;
        List<String> choices;
    }
}
