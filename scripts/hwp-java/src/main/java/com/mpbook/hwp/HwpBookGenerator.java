package com.mpbook.hwp;

import com.google.gson.Gson;
import com.google.gson.annotations.SerializedName;
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlEquation;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.GsoHeaderProperty;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.HeightCriterion;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.HorzRelTo;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.ObjectNumberSort;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.RelativeArrange;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.TextFlowMethod;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.TextHorzArrange;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.VertRelTo;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.WidthCriterion;
import kr.dogfoot.hwplib.object.bodytext.control.gso.ControlRectangle;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControlType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.ShapeComponentNormal;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.LineArrowShape;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.LineArrowSize;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.LineEndShape;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.LineInfo;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.LineType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.OutlineStyle;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowInfo;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.ShapeComponentRectangle;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharControlExtend;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.ParaText;
import kr.dogfoot.hwplib.object.docinfo.BinData;
import kr.dogfoot.hwplib.object.docinfo.ParaShape;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataCompress;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataState;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.FillInfo;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.ImageFill;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.ImageFillType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureEffect;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.Reader;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class HwpBookGenerator {
    private static final Gson GSON = new Gson();
    private static final String DEFAULT_EQUATION_FONT_NAME = "HancomEQN";
    private static final boolean ENABLE_EXPLICIT_EQUATION_FONT = true;
    private static final char SQUARE_PLACEHOLDER_CHAR = '\u25A1';
    private static final int TEMPLATE_EQ_WIDTH = 10105;
    private static final int TEMPLATE_EQ_HEIGHT = 1163;
    private static final double MM_PER_INCH = 25.4;
    private static final int DEFAULT_PARA_SHAPE_ID = 3;
    private static final short DEFAULT_STYLE_ID = 0;
    private static final int PAGE_TEXT_WIDTH_MM = 160;
    private static final int DEFAULT_VISUAL_WIDTH_MM = 56;
    private static final int MAX_VISUAL_WIDTH_MM = 150;
    private static final BinDataCompress IMAGE_COMPRESS_METHOD = BinDataCompress.ByStorageDefault;
    private static final String MODE_INLINE = "inline";
    private static final String MODE_RIGHT_IMAGE = "right_image";
    private static final Map<String, VisualLayoutSpec> VISUAL_LAYOUTS = createVisualLayouts();

    private static final Pattern INLINE_FORMULA_PATTERN = Pattern.compile("\\\\\\((.+?)\\\\\\)");
    private static final Pattern COMMAND_PATTERN = Pattern.compile("\\\\([A-Za-z]+)");
    private static final Pattern WORD_PATTERN = Pattern.compile("[A-Za-z]+");
    private static final Pattern FRAC_SIMPLE_PATTERN = Pattern.compile("\\\\frac\\s*\\{([^{}]+)}\\s*\\{([^{}]+)}");
    private static final Pattern SQRT_SIMPLE_PATTERN = Pattern.compile("\\\\sqrt\\s*\\{([^{}]+)}");
    private static final Pattern OVERLINE_SIMPLE_PATTERN = Pattern.compile("\\\\overline\\s*\\{([^{}]+)}");
    private static final Pattern MATHRM_PATTERN = Pattern.compile("\\\\mathrm\\s*\\{([^{}]+)}");
    private static final Pattern TEXT_PATTERN = Pattern.compile("\\\\text\\s*\\{([^{}]+)}");
    private static final Pattern LOG_SUBSCRIPT_PATTERN = Pattern.compile("\\blog\\s*_\\s*([0-9A-Za-z]+)");
    private static final Pattern IMPLICIT_MUL_SQRT_PATTERN = Pattern.compile("([0-9A-Za-z}\\)])\\s*sqrt\\{");
    private static final Pattern HANGUL_PATTERN = Pattern.compile("[\\uAC00-\\uD7A3\\u3131-\\u314E\\u314F-\\u3163]");
    private static final Set<String> SUPPORTED_LATEX_COMMANDS = new HashSet<>(Arrays.asList(
            "frac", "sqrt", "mathrm", "times", "cdot", "div",
            "leq", "geq", "le", "ge", "neq", "pm", "mp",
            "alpha", "beta", "gamma", "theta", "pi",
            "sin", "cos", "tan", "log", "left", "right", "circ",
            "text", "square",
            "angle", "overline"
    ));
    private static final Set<String> ALLOWED_SCRIPT_WORDS = new HashSet<>(Arrays.asList(
            "sin", "cos", "tan", "log", "sqrt", "over", "times",
            "leq", "geq", "neq", "pm", "mp",
            "alpha", "beta", "gamma", "theta", "pi", "angle", "circ",
            "cm", "mm", "km", "m"
    ));

    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);

        Path inputJson = Paths.get(parsed.inputJsonPath).toAbsolutePath().normalize();
        Path outputHwp = Paths.get(parsed.outputHwpPath).toAbsolutePath().normalize();
        Path visualDir = Paths.get(parsed.visualDirPath).toAbsolutePath().normalize();

        QuestionPayload payload = loadPayload(inputJson);
        HWPFile hwpFile = buildBook(payload, visualDir);

        Files.createDirectories(outputHwp.getParent());
        HWPWriter.toFile(hwpFile, outputHwp.toString());

        System.out.println("Generated HWP (safe+equation mode): " + outputHwp);
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

    private static HWPFile buildBook(QuestionPayload payload, Path visualDir) throws Exception {
        HWPFile hwpFile = BlankFileMaker.make();
        Section section = hwpFile.getBodyText().getSectionList().get(0);
        int questionStartParaShapeId = ensureQuestionStartParaShapeId(hwpFile, DEFAULT_PARA_SHAPE_ID);

        int index = 0;
        for (Question q : safeQuestions(payload.questions)) {
            index++;
            String qid = valueOrFallback(q.id, String.format("Q%02d", index));
            String qtype = valueOrFallback(q.type, "unknown");
            String title = valueOrFallback(q.title, "");
            int titleParaShapeId = index > 1 ? questionStartParaShapeId : DEFAULT_PARA_SHAPE_ID;

            if (title.isEmpty()) {
                addParagraph(section, qid + " [" + qtype + "]", titleParaShapeId);
            } else {
                addParagraph(section, qid + " [" + qtype + "] " + title, titleParaShapeId);
            }
            addSpacer(section, 1);

            List<String> bodyLines = safeLines(q.body);
            List<VisualItem> visualItems = resolveVisualItems(
                    visualDir,
                    qid,
                    q.sourceImage,
                    VISUAL_LAYOUTS.get(q.sourceImage)
            );
            Map<Integer, List<VisualItem>> visualAfterMap = groupVisualsByInsertIndex(visualItems);

            for (int bodyIdx = 0; bodyIdx < bodyLines.size(); bodyIdx++) {
                String bodyLine = bodyLines.get(bodyIdx);
                if (bodyLine == null || bodyLine.isBlank()) {
                    addParagraph(section, "");
                } else {
                    addRichParagraph(section, bodyLine);
                }

                List<VisualItem> visuals = visualAfterMap.getOrDefault(bodyIdx, Collections.emptyList());
                if (!visuals.isEmpty()) {
                    addSpacer(section, 1);
                }
                for (VisualItem visual : visuals) {
                    if (!addVisualImage(hwpFile, section, visual)) {
                        addParagraph(section, "[visual] " + visual.path.getFileName());
                    }
                }
                if (!visuals.isEmpty()) {
                    addSpacer(section, 1);
                }
            }

            if (bodyLines.isEmpty()) {
                List<VisualItem> visuals = visualAfterMap.getOrDefault(0, Collections.emptyList());
                if (!visuals.isEmpty()) {
                    addSpacer(section, 1);
                }
                for (VisualItem visual : visuals) {
                    if (!addVisualImage(hwpFile, section, visual)) {
                        addParagraph(section, "[visual] " + visual.path.getFileName());
                    }
                }
                if (!visuals.isEmpty()) {
                    addSpacer(section, 1);
                }
            }

            List<String> choices = safeLines(q.choices);
            if (!choices.isEmpty()) {
                addParagraph(section, "Choices:");
                int cIdx = 1;
                for (String c : choices) {
                    addRichParagraph(section, "(" + cIdx + ") " + (c == null ? "" : c));
                    cIdx++;
                }
            }

            addSpacer(section, 1);
        }

        return hwpFile;
    }

    private static void addRichParagraph(Section section, String input) throws UnsupportedEncodingException {
        addRichParagraph(section, input, DEFAULT_PARA_SHAPE_ID, DEFAULT_STYLE_ID);
    }

    private static void addRichParagraph(Section section, String input, int paraShapeId) throws UnsupportedEncodingException {
        addRichParagraph(section, input, paraShapeId, DEFAULT_STYLE_ID);
    }

    private static void addRichParagraph(Section section, String input, int paraShapeId, short styleId) throws UnsupportedEncodingException {
        String normalized = normalizeLine(input);
        Paragraph paragraph = section.addNewParagraph();
        initParagraphHeader(paragraph, paraShapeId, styleId);
        paragraph.createText();
        paragraph.createCharShape();
        paragraph.getCharShape().addParaCharShape(0, 0);

        ParaText paraText = paragraph.getText();
        Matcher matcher = INLINE_FORMULA_PATTERN.matcher(normalized);
        int cursor = 0;
        boolean wroteAny = false;

        while (matcher.find()) {
            String plainPart = normalized.substring(cursor, matcher.start());
            if (!plainPart.isEmpty()) {
                appendInlinePlainText(paraText, plainPart);
                wroteAny = true;
            }

            String latex = matcher.group(1) == null ? "" : matcher.group(1).trim();
            String script = latexToHwpEquationScript(latex);
            if (script != null && addEquationInline(paragraph, script)) {
                wroteAny = true;
            } else if (!latex.isEmpty()) {
                appendInlinePlainText(paraText, latex);
                wroteAny = true;
            }

            cursor = matcher.end();
        }

        String tail = normalized.substring(cursor);
        if (!tail.isEmpty()) {
            appendInlinePlainText(paraText, tail);
            wroteAny = true;
        }

        if (!wroteAny) {
            appendInlinePlainText(paraText, "");
        }
        ensureParaBreak(paraText);
    }

    private static void addParagraph(Section section, String text) throws UnsupportedEncodingException {
        addParagraph(section, text, DEFAULT_PARA_SHAPE_ID, DEFAULT_STYLE_ID);
    }

    private static void addParagraph(Section section, String text, int paraShapeId) throws UnsupportedEncodingException {
        addParagraph(section, text, paraShapeId, DEFAULT_STYLE_ID);
    }

    private static void addParagraph(Section section, String text, int paraShapeId, short styleId) throws UnsupportedEncodingException {
        Paragraph paragraph = section.addNewParagraph();
        initParagraphHeader(paragraph, paraShapeId, styleId);

        paragraph.createText();
        paragraph.getText().addString(text == null ? "" : text);

        paragraph.createCharShape();
        paragraph.getCharShape().addParaCharShape(0, 0);
    }

    private static void initParagraphHeader(Paragraph paragraph) {
        initParagraphHeader(paragraph, DEFAULT_PARA_SHAPE_ID, DEFAULT_STYLE_ID);
    }

    private static void initParagraphHeader(Paragraph paragraph, int paraShapeId, short styleId) {
        paragraph.getHeader().setParaShapeId(Math.max(0, paraShapeId));
        paragraph.getHeader().setStyleId(styleId);
        paragraph.getHeader().getDivideSort().setValue((short) 0);
        paragraph.getHeader().setInstanceID(0);
        paragraph.getHeader().setIsMergedByTrack(0);
    }

    private static void addSpacer(Section section, int count) throws UnsupportedEncodingException {
        int safeCount = Math.max(0, count);
        for (int i = 0; i < safeCount; i++) {
            addParagraph(section, "");
        }
    }

    private static int ensureQuestionStartParaShapeId(HWPFile hwpFile, int baseParaShapeId) {
        List<ParaShape> paraShapes = hwpFile.getDocInfo().getParaShapeList();
        if (paraShapes == null || paraShapes.isEmpty()) {
            return Math.max(0, baseParaShapeId);
        }

        int sourceId = baseParaShapeId;
        if (sourceId < 0 || sourceId >= paraShapes.size()) {
            sourceId = 0;
        }

        ParaShape pageBreakShape = paraShapes.get(sourceId).clone();
        pageBreakShape.getProperty1().setSplitPageBeforePara(true);
        paraShapes.add(pageBreakShape);
        return paraShapes.size() - 1;
    }

    private static boolean addEquationInline(Paragraph paragraph, String script) {
        String normalizedScript = script == null ? "" : script.trim();
        if (!isSafeEquationScript(normalizedScript)) {
            return false;
        }

        try {
            ParaText paraText = paragraph.getText();
            addExtendCharForEquation(paraText);

            ControlEquation equation = (ControlEquation) paragraph.addNewControl(ControlType.Equation);
            setEquationHeader(equation, normalizedScript);
            setEquationScript(equation, normalizedScript);
            return true;
        } catch (Exception ignored) {
            return false;
        }
    }

    private static void appendInlinePlainText(ParaText paraText, String text) {
        if (text == null) {
            return;
        }
        for (int i = 0; i < text.length(); i++) {
            paraText.addNewNormalChar().setCode((short) text.charAt(i));
        }
    }

    private static void ensureParaBreak(ParaText paraText) {
        paraText.addNewNormalChar().setCode((short) 0x000d);
    }

    private static void addExtendCharForEquation(ParaText paraText) throws Exception {
        HWPCharControlExtend chExtend = paraText.addNewExtendControlChar();
        chExtend.setCode((short) 0x000b);
        byte[] addition = new byte[12];
        addition[3] = 'e';
        addition[2] = 'q';
        addition[1] = 'e';
        addition[0] = 'd';
        chExtend.setAddition(addition);
    }

    private static void setEquationHeader(ControlEquation equation, String script) {
        CtrlHeaderGso hdr = equation.getHeader();
        GsoHeaderProperty prop = hdr.getProperty();
        prop.setLikeWord(true);
        prop.setApplyLineSpace(false);
        prop.setVertRelTo(VertRelTo.Para);
        prop.setVertRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setHorzRelTo(HorzRelTo.Para);
        prop.setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setVertRelToParaLimit(true);
        prop.setAllowOverlap(false);
        prop.setWidthCriterion(WidthCriterion.Absolute);
        prop.setHeightCriterion(HeightCriterion.Absolute);
        prop.setProtectSize(false);
        prop.setTextFlowMethod(TextFlowMethod.TakePlace);
        prop.setTextHorzArrange(TextHorzArrange.BothSides);
        prop.setObjectNumberSort(ObjectNumberSort.Equation);

        hdr.setyOffset(0);
        hdr.setxOffset(0);
        hdr.setWidth(TEMPLATE_EQ_WIDTH);
        hdr.setHeight(TEMPLATE_EQ_HEIGHT);
        hdr.setzOrder(0);
        hdr.setOutterMarginLeft(0);
        hdr.setOutterMarginRight(0);
        hdr.setOutterMarginTop(0);
        hdr.setOutterMarginBottom(0);
        hdr.setInstanceId(0);
        hdr.setPreventPageDivide(false);
        hdr.getExplanation().setBytes(null);
    }

    private static void setEquationScript(ControlEquation equation, String script) {
        equation.getEQEdit().setProperty(0);
        equation.getEQEdit().getScript().fromUTF16LEString(script);
        equation.getEQEdit().setLetterSize(1000);
        equation.getEQEdit().getLetterColor().setValue(0);
        equation.getEQEdit().setBaseLine((short) 89);
        equation.getEQEdit().setUnknown(0);
        equation.getEQEdit().getVersionInfo().fromUTF16LEString("Equation Version 60");
        if (ENABLE_EXPLICIT_EQUATION_FONT) {
            equation.getEQEdit().getFontName().fromUTF16LEString(DEFAULT_EQUATION_FONT_NAME);
        }
    }

    private static String normalizeLine(String input) {
        if (input == null) {
            return "";
        }
        return input.replace("\r\n", " ").replace('\r', ' ').replace('\n', ' ').replace('\t', ' ');
    }

    private static String latexToHwpEquationScript(String latex) {
        if (latex == null) {
            return null;
        }
        String raw = latex.trim();
        if (raw.isEmpty()) {
            return null;
        }
        if (!hasOnlySupportedLatexCommands(raw)) {
            return null;
        }

        String script = raw;
        script = script.replace("\r", " ").replace("\n", " ");
        script = script.replace("\\left", "").replace("\\right", "");
        script = script.replace("\\,", " ");
        script = stripTextCommands(script);

        script = convertFracSqrtRecursively(script);
        script = replacePatternIterative(FRAC_SIMPLE_PATTERN, script, "{$1} over {$2}");
        script = replacePatternIterative(SQRT_SIMPLE_PATTERN, script, "sqrt{$1}");
        script = replacePatternIterative(OVERLINE_SIMPLE_PATTERN, script, "$1");
        script = replacePatternIterative(MATHRM_PATTERN, script, "$1");
        script = replacePatternIterative(TEXT_PATTERN, script, "$1");

        script = script.replace("\\times", " times ");
        script = script.replace("\\cdot", " times ");
        script = script.replace("\\div", " / ");
        script = script.replace("\\leq", " leq ");
        script = script.replace("\\geq", " geq ");
        script = script.replace("\\le", " leq ");
        script = script.replace("\\ge", " geq ");
        script = script.replace("\\neq", " neq ");
        script = script.replace("\\pm", " pm ");
        script = script.replace("\\mp", " mp ");
        script = script.replace("\\alpha", " alpha ");
        script = script.replace("\\beta", " beta ");
        script = script.replace("\\gamma", " gamma ");
        script = script.replace("\\theta", " theta ");
        script = script.replace("\\pi", " pi ");
        script = script.replace("\\sin", " sin ");
        script = script.replace("\\cos", " cos ");
        script = script.replace("\\tan", " tan ");
        script = script.replace("\\log", " log ");
        script = script.replace("\\angle", " angle ");
        script = script.replace("\\square", String.valueOf(SQUARE_PLACEHOLDER_CHAR));
        script = script.replace("\\circ", "circ");
        script = script.replace("^circ", "^{circ}");

        script = script.replace("\\", "");
        script = LOG_SUBSCRIPT_PATTERN.matcher(script).replaceAll("log_$1");
        script = IMPLICIT_MUL_SQRT_PATTERN.matcher(script).replaceAll("$1*sqrt{");
        script = script.replaceAll("\\s+", " ").trim();
        return isSafeEquationScript(script) ? script : null;
    }

    private static String convertFracSqrtRecursively(String input) {
        String prev = input;
        while (true) {
            String next = convertOnePassFracSqrt(prev);
            if (next.equals(prev)) {
                return next;
            }
            prev = next;
        }
    }

    private static String convertOnePassFracSqrt(String input) {
        StringBuilder out = new StringBuilder();
        int i = 0;
        while (i < input.length()) {
            if (input.startsWith("\\frac", i)) {
                int p = i + 5;
                p = skipSpaces(input, p);
                ParsedBrace a = parseBrace(input, p);
                if (a != null) {
                    int q = skipSpaces(input, a.end + 1);
                    ParsedBrace b = parseBrace(input, q);
                    if (b != null) {
                        out.append("{")
                                .append(convertFracSqrtRecursively(a.content))
                                .append("} over {")
                                .append(convertFracSqrtRecursively(b.content))
                                .append("}");
                        i = b.end + 1;
                        continue;
                    }
                }
            }

            if (input.startsWith("\\sqrt", i)) {
                int p = i + 5;
                p = skipSpaces(input, p);
                ParsedBrace a = parseBrace(input, p);
                if (a != null) {
                    out.append("sqrt{")
                            .append(convertFracSqrtRecursively(a.content))
                            .append("}");
                    i = a.end + 1;
                    continue;
                }
            }

            out.append(input.charAt(i));
            i++;
        }
        return out.toString();
    }

    private static int skipSpaces(String input, int index) {
        int p = index;
        while (p < input.length() && Character.isWhitespace(input.charAt(p))) {
            p++;
        }
        return p;
    }

    private static ParsedBrace parseBrace(String input, int start) {
        if (start >= input.length() || input.charAt(start) != '{') {
            return null;
        }

        int depth = 0;
        StringBuilder content = new StringBuilder();
        for (int i = start; i < input.length(); i++) {
            char c = input.charAt(i);
            if (c == '{') {
                depth++;
                if (depth > 1) {
                    content.append(c);
                }
            } else if (c == '}') {
                depth--;
                if (depth == 0) {
                    return new ParsedBrace(content.toString(), i);
                }
                if (depth > 0) {
                    content.append(c);
                }
            } else if (depth > 0) {
                content.append(c);
            }
        }

        return null;
    }

    private static String replacePatternIterative(Pattern pattern, String input, String replacement) {
        String prev = input;
        while (true) {
            String next = pattern.matcher(prev).replaceAll(replacement);
            if (next.equals(prev)) {
                return next;
            }
            prev = next;
        }
    }

    private static String stripTextCommands(String input) {
        String prev = input;
        while (true) {
            Matcher matcher = TEXT_PATTERN.matcher(prev);
            if (!matcher.find()) {
                return prev;
            }
            matcher.reset();

            StringBuffer sb = new StringBuffer();
            while (matcher.find()) {
                String inner = matcher.group(1) == null ? "" : matcher.group(1);
                String replacement = HANGUL_PATTERN.matcher(inner).find() ? "" : inner;
                matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
            }
            matcher.appendTail(sb);
            String next = sb.toString();
            if (next.equals(prev)) {
                return next;
            }
            prev = next;
        }
    }

    private static boolean isSafeEquationScript(String script) {
        if (script == null || script.isBlank()) {
            return false;
        }
        if (script.length() > 180) {
            return false;
        }
        if (HANGUL_PATTERN.matcher(script).find()) {
            return false;
        }

        int brace = 0;
        int paren = 0;
        for (int i = 0; i < script.length(); i++) {
            char ch = script.charAt(i);
            if ((ch < 0x20 || ch > 0x7E) && ch != SQUARE_PLACEHOLDER_CHAR) {
                return false;
            }
            if (ch == '[' || ch == ']') {
                return false;
            }
            if (ch == '{') {
                brace++;
            } else if (ch == '}') {
                brace--;
                if (brace < 0) {
                    return false;
                }
            } else if (ch == '(') {
                paren++;
            } else if (ch == ')') {
                paren--;
                if (paren < 0) {
                    return false;
                }
            }
        }
        if (!hasBalancedPipes(script)) {
            return false;
        }

        if (brace != 0 || paren != 0) {
            return false;
        }

        Matcher wordMatcher = WORD_PATTERN.matcher(script);
        while (wordMatcher.find()) {
            String token = wordMatcher.group();
            if (!isAllowedScriptWordToken(token)) {
                return false;
            }
        }

        return true;
    }

    private static boolean hasOnlySupportedLatexCommands(String latex) {
        Matcher matcher = COMMAND_PATTERN.matcher(latex);
        while (matcher.find()) {
            String cmd = matcher.group(1);
            if (!SUPPORTED_LATEX_COMMANDS.contains(cmd)) {
                return false;
            }
        }
        return true;
    }

    private static boolean isAllowedScriptWordToken(String token) {
        String lower = token.toLowerCase(Locale.ROOT);
        if (ALLOWED_SCRIPT_WORDS.contains(lower)) {
            return true;
        }
        if (token.length() == 1 && Character.isLowerCase(token.charAt(0))) {
            return true;
        }
        if (token.equals(token.toUpperCase(Locale.ROOT))) {
            return true;
        }
        return token.length() <= 6 && !token.equals(token.toLowerCase(Locale.ROOT));
    }

    private static boolean hasBalancedPipes(String script) {
        int pipes = 0;
        for (int i = 0; i < script.length(); i++) {
            if (script.charAt(i) == '|') {
                pipes++;
            }
        }
        return pipes % 2 == 0;
    }

    private static Map<Integer, List<VisualItem>> groupVisualsByInsertIndex(List<VisualItem> visualItems) {
        if (visualItems == null || visualItems.isEmpty()) {
            return Collections.emptyMap();
        }
        Map<Integer, List<VisualItem>> grouped = new HashMap<>();
        for (VisualItem visual : visualItems) {
            grouped.computeIfAbsent(Math.max(0, visual.insertAfterBodyIndex), k -> new ArrayList<>()).add(visual);
        }
        return grouped;
    }

    private static List<VisualItem> resolveVisualItems(
            Path visualDir,
            String qid,
            String sourceImage,
            VisualLayoutSpec layoutSpec
    ) {
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

        List<VisualItem> items = new ArrayList<>();
        String mode = layoutSpec == null ? MODE_INLINE : layoutSpec.mode;
        List<VisualItemSpec> specs = layoutSpec == null ? Collections.emptyList() : layoutSpec.items;

        for (int i = 0; i < matches.size(); i++) {
            Path p = matches.get(i);
            VisualItemSpec spec = i < specs.size() ? specs.get(i) : null;

            double widthIn;
            String align;
            int insertAfterBodyIndex;

            if (spec != null) {
                widthIn = spec.widthIn;
                align = spec.align;
                insertAfterBodyIndex = spec.insertAfterBodyIndex;
            } else {
                widthIn = MODE_RIGHT_IMAGE.equals(mode) ? 1.9 : 2.6;
                align = MODE_RIGHT_IMAGE.equals(mode) ? "right" : "center";
                insertAfterBodyIndex = 0;
            }

            items.add(new VisualItem(
                    p,
                    sourceImage == null ? "" : sourceImage,
                    widthIn,
                    align,
                    insertAfterBodyIndex
            ));
        }

        return items;
    }

    private static boolean addVisualImage(HWPFile hwpFile, Section section, VisualItem visual) {
        if (visual == null || visual.path == null || !Files.exists(visual.path)) {
            return false;
        }

        try {
            int[] mm = visualSizeMm(visual.path, visual.widthIn);
            int widthMm = mm[0];
            int heightMm = mm[1];
            int xOffsetMm = visualXOffsetMm(visual.align, widthMm);
            int binDataId = addBinDataForImage(hwpFile, visual.path);

            Paragraph paragraph = section.addNewParagraph();
            initParagraphHeader(paragraph);
            paragraph.createText();
            paragraph.getText().addExtendCharForGSO();
            paragraph.createCharShape();
            paragraph.getCharShape().addParaCharShape(0, 0);

            ControlRectangle rectangle = (ControlRectangle) paragraph.addNewGsoControl(GsoControlType.Rectangle);
            setVisualCtrlHeader(rectangle, xOffsetMm, widthMm, heightMm);
            setVisualShape(rectangle, widthMm, heightMm, binDataId);
            return true;
        } catch (Exception ignored) {
            return false;
        }
    }

    private static int[] visualSizeMm(Path imagePath, double widthIn) throws IOException {
        int widthMm = (int) Math.round(widthIn * MM_PER_INCH);
        widthMm = Math.max(20, Math.min(MAX_VISUAL_WIDTH_MM, widthMm));

        BufferedImage image = ImageIO.read(imagePath.toFile());
        if (image == null || image.getWidth() <= 0 || image.getHeight() <= 0) {
            return new int[]{widthMm, Math.max(14, (int) Math.round(widthMm * 0.66))};
        }

        int heightMm = (int) Math.round((double) widthMm * (double) image.getHeight() / (double) image.getWidth());
        heightMm = Math.max(12, heightMm);
        return new int[]{widthMm, heightMm};
    }

    private static int visualXOffsetMm(String align, int widthMm) {
        int max = Math.max(0, PAGE_TEXT_WIDTH_MM - widthMm);
        if ("right".equalsIgnoreCase(align)) {
            return max;
        }
        if ("left".equalsIgnoreCase(align)) {
            return 0;
        }
        return Math.max(0, max / 2);
    }

    private static int addBinDataForImage(HWPFile hwpFile, Path imagePath) throws IOException {
        int streamIndex = hwpFile.getBinData().getEmbeddedBinaryDataList().size() + 1;
        String ext = normalizeImageExtension(imagePath);
        String streamName = "Bin" + String.format("%04X", streamIndex) + "." + ext;
        byte[] fileBinary = Files.readAllBytes(imagePath);

        hwpFile.getBinData().addNewEmbeddedBinaryData(streamName, fileBinary, IMAGE_COMPRESS_METHOD);

        BinData bd = new BinData();
        bd.getProperty().setType(BinDataType.Embedding);
        bd.getProperty().setCompress(IMAGE_COMPRESS_METHOD);
        bd.getProperty().setState(BinDataState.NotAccess);
        bd.setBinDataID(streamIndex);
        bd.setExtensionForEmbedding(ext);
        hwpFile.getDocInfo().getBinDataList().add(bd);
        return hwpFile.getDocInfo().getBinDataList().size();
    }

    private static String normalizeImageExtension(Path imagePath) {
        String name = imagePath.getFileName().toString();
        int dot = name.lastIndexOf('.');
        String ext = dot >= 0 ? name.substring(dot + 1).toLowerCase(Locale.ROOT) : "png";
        if ("jpeg".equals(ext)) {
            return "jpg";
        }
        if ("png".equals(ext) || "jpg".equals(ext) || "bmp".equals(ext) || "gif".equals(ext)) {
            return ext;
        }
        return "png";
    }

    private static void setVisualCtrlHeader(ControlRectangle rectangle, int xOffsetMm, int widthMm, int heightMm) {
        CtrlHeaderGso hdr = rectangle.getHeader();
        GsoHeaderProperty prop = hdr.getProperty();
        prop.setLikeWord(false);
        prop.setApplyLineSpace(false);
        prop.setVertRelTo(VertRelTo.Para);
        prop.setVertRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setHorzRelTo(HorzRelTo.Para);
        prop.setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setVertRelToParaLimit(true);
        prop.setAllowOverlap(false);
        prop.setWidthCriterion(WidthCriterion.Absolute);
        prop.setHeightCriterion(HeightCriterion.Absolute);
        prop.setProtectSize(false);
        prop.setTextFlowMethod(TextFlowMethod.FitWithText);
        prop.setTextHorzArrange(TextHorzArrange.BothSides);
        prop.setObjectNumberSort(ObjectNumberSort.Figure);

        hdr.setyOffset(fromMMForGso(0));
        hdr.setxOffset(fromMMForGso(Math.max(0, xOffsetMm)));
        hdr.setWidth(fromMMForGso(widthMm));
        hdr.setHeight(fromMMForGso(heightMm));
        hdr.setzOrder(0);
        hdr.setOutterMarginLeft(0);
        hdr.setOutterMarginRight(0);
        hdr.setOutterMarginTop(0);
        hdr.setOutterMarginBottom(0);
        hdr.setInstanceId(0);
        hdr.setPreventPageDivide(false);
        hdr.getExplanation().setBytes(null);
    }

    private static void setVisualShape(ControlRectangle rectangle, int widthMm, int heightMm, int binDataId) {
        ShapeComponentNormal sc = (ShapeComponentNormal) rectangle.getShapeComponent();
        sc.getProperty().setRotateWithImage(true);
        sc.setOffsetX(0);
        sc.setOffsetY(0);
        sc.setGroupingCount(0);
        sc.setLocalFileVersion(1);
        sc.setWidthAtCreate(fromMMForGso(widthMm));
        sc.setHeightAtCreate(fromMMForGso(heightMm));
        sc.setWidthAtCurrent(fromMMForGso(widthMm));
        sc.setHeightAtCurrent(fromMMForGso(heightMm));
        sc.setRotateAngle(0);
        sc.setRotateXCenter(fromMMForGso(widthMm / 2));
        sc.setRotateYCenter(fromMMForGso(heightMm / 2));

        sc.createLineInfo();
        LineInfo li = sc.getLineInfo();
        li.getProperty().setLineEndShape(LineEndShape.Flat);
        li.getProperty().setStartArrowShape(LineArrowShape.None);
        li.getProperty().setStartArrowSize(LineArrowSize.MiddleMiddle);
        li.getProperty().setEndArrowShape(LineArrowShape.None);
        li.getProperty().setEndArrowSize(LineArrowSize.MiddleMiddle);
        li.getProperty().setFillStartArrow(true);
        li.getProperty().setFillEndArrow(true);
        li.getProperty().setLineType(LineType.None);
        li.setOutlineStyle(OutlineStyle.Normal);
        li.setThickness(0);
        li.getColor().setValue(0);

        sc.createFillInfo();
        FillInfo fi = sc.getFillInfo();
        fi.getType().setPatternFill(false);
        fi.getType().setImageFill(true);
        fi.getType().setGradientFill(false);
        fi.createImageFill();
        ImageFill imgF = fi.getImageFill();
        imgF.setImageFillType(ImageFillType.FitSize);
        imgF.getPictureInfo().setBrightness((byte) 0);
        imgF.getPictureInfo().setContrast((byte) 0);
        imgF.getPictureInfo().setEffect(PictureEffect.RealPicture);
        imgF.getPictureInfo().setBinItemID(binDataId);

        sc.createShadowInfo();
        ShadowInfo si = sc.getShadowInfo();
        si.setType(ShadowType.None);
        si.getColor().setValue(0xc4c4c4);
        si.setOffsetX(283);
        si.setOffsetY(283);
        si.setTransparent((short) 0);
        sc.setMatrixsNormal();

        ShapeComponentRectangle scr = rectangle.getShapeComponentRectangle();
        scr.setRoundRate((byte) 0);
        scr.setX1(0);
        scr.setY1(0);
        scr.setX2(fromMMForGso(widthMm));
        scr.setY2(0);
        scr.setX3(fromMMForGso(widthMm));
        scr.setY3(fromMMForGso(heightMm));
        scr.setX4(0);
        scr.setY4(fromMMForGso(heightMm));
    }

    private static int fromMMForGso(int mm) {
        if (mm == 0) {
            return 1;
        }
        return (int) ((double) mm * 72000.0f / 254.0f + 0.5f);
    }

    private static Map<String, VisualLayoutSpec> createVisualLayouts() {
        Map<String, VisualLayoutSpec> map = new HashMap<>();

        map.put("sample_input/sample_blank_1.png", layout(MODE_RIGHT_IMAGE,
                item(1.9, "right", 0)));
        map.put("sample_input/sample_blank_2.png", layout(MODE_RIGHT_IMAGE,
                item(1.9, "right", 0)));
        map.put("sample_input/sample_blank_3.png", layout(MODE_INLINE,
                item(2.8, "left", 0),
                item(2.8, "left", 0)));
        map.put("sample_input/sample_graph_1.png", layout(MODE_INLINE,
                item(2.9, "center", 0)));
        map.put("sample_input/sample_graph_2.png", layout(MODE_RIGHT_IMAGE,
                item(1.75, "right", 0)));
        map.put("sample_input/sample_graph_3.png", layout(MODE_RIGHT_IMAGE,
                item(1.9, "right", 0)));
        map.put("sample_input/sample_graph_4.png", layout(MODE_RIGHT_IMAGE,
                item(1.95, "right", 0)));
        map.put("sample_input/sample_graph_5.png", layout(MODE_RIGHT_IMAGE,
                item(1.95, "right", 0)));
        map.put("sample_input/sample_table_1.png", layout(MODE_INLINE,
                item(5.8, "center", 0)));
        map.put("sample_input/sample_table_2.png", layout(MODE_INLINE,
                item(5.6, "center", 0)));

        return map;
    }

    private static VisualLayoutSpec layout(String mode, VisualItemSpec... items) {
        List<VisualItemSpec> specs = new ArrayList<>();
        if (items != null) {
            Collections.addAll(specs, items);
        }
        return new VisualLayoutSpec(mode, specs);
    }

    private static VisualItemSpec item(double widthIn, String align, int insertAfterBodyIndex) {
        return new VisualItemSpec(widthIn, align, insertAfterBodyIndex);
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

    private static final class ParsedBrace {
        private final String content;
        private final int end;

        private ParsedBrace(String content, int end) {
            this.content = content;
            this.end = end;
        }
    }

    private static final class VisualLayoutSpec {
        private final String mode;
        private final List<VisualItemSpec> items;

        private VisualLayoutSpec(String mode, List<VisualItemSpec> items) {
            this.mode = mode == null ? MODE_INLINE : mode;
            this.items = items == null ? Collections.emptyList() : items;
        }
    }

    private static final class VisualItemSpec {
        private final double widthIn;
        private final String align;
        private final int insertAfterBodyIndex;

        private VisualItemSpec(double widthIn, String align, int insertAfterBodyIndex) {
            this.widthIn = widthIn;
            this.align = align == null ? "center" : align;
            this.insertAfterBodyIndex = Math.max(0, insertAfterBodyIndex);
        }
    }

    private static final class VisualItem {
        private final Path path;
        @SuppressWarnings("unused")
        private final String sourceImage;
        private final double widthIn;
        private final String align;
        private final int insertAfterBodyIndex;

        private VisualItem(Path path, String sourceImage, double widthIn, String align, int insertAfterBodyIndex) {
            this.path = path;
            this.sourceImage = sourceImage;
            this.widthIn = widthIn;
            this.align = align == null ? "center" : align;
            this.insertAfterBodyIndex = Math.max(0, insertAfterBodyIndex);
        }
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
