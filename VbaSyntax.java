import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/** Token classes, keyword tables, completion data and indentation rules for the macro language. */
public final class VbaSyntax {
    private VbaSyntax() {
    }

    public static final int PLAIN = 0, KEYWORD = 1, STRING = 2, COMMENT = 3, NUMBER = 4, BUILTIN = 5, OBJECT = 6, TYPE = 7,
            CONSTANT = 8, MEMBER = 9, OPERATOR = 10, KINDS = 11;

    /** One coloured piece of a line. */
    public static final class Span {
        public final int start, end, kind;
        /** Preferred spelling (keywords, types, constants) or null. */
        public final String canonical;

        Span(int start, int end, int kind, String canonical) {
            this.start = start;
            this.end = end;
            this.kind = kind;
            this.canonical = canonical;
        }
    }

    private static final String[] KEYWORDS = {
            "Sub", "Function", "End", "If", "Then", "Else", "ElseIf", "For", "To", "Step", "Next", "Each", "In", "Do", "While",
            "Until", "Loop", "Wend", "Select", "Case", "With", "Dim", "ReDim", "Preserve", "Const", "As", "Set", "Call", "Exit",
            "On", "Error", "Resume", "GoTo", "Not", "And", "Or", "Xor", "Mod", "Like", "ByVal", "ByRef", "Optional", "Public",
            "Private", "Static", "Option", "Explicit", "Return", "True", "False", "Nothing", "Empty", "Null", "New", "Is", "Let",
            "Stop", "Randomize", "Type", "Enum", "Friend", "ParamArray", "Declare", "Attribute"};
    private static final String[] TYPES = {"Integer", "Long", "Double", "Single", "String", "Boolean", "Variant", "Object",
            "Byte", "Currency", "Date", "LongLong"};
    private static final String[] BUILTINS = {
            "Abs", "Array", "Asc", "Atn", "CBool", "CDbl", "CInt", "CLng", "CStr", "Chr", "Cos", "DateAdd", "DateDiff", "DateSerial",
            "Day", "Evaluate", "Exp", "Fix", "Format", "FormatCurrency", "FormatNumber", "FormatPercent", "InStr", "InStrRev",
            "InputBox", "Int", "IsArray", "IsDate", "IsEmpty", "IsNumeric", "Join", "LBound", "LCase", "LTrim", "Left", "Len", "Log",
            "MsgBox", "Mid", "Month", "Now", "RGB", "RTrim", "Replace", "Right", "Rnd", "Round", "Sgn", "Sin", "Space", "Split", "Sqr",
            "Str", "StrComp", "StrReverse", "Tan", "Time", "Timer", "Trim", "TypeName", "UBound", "UCase", "Val", "Weekday", "Year"};
    private static final String[] OBJECTS = {"Range", "Cells", "Rows", "Columns", "ActiveCell", "Selection", "ActiveSheet",
            "Worksheets", "Sheets", "ThisWorkbook", "ActiveWorkbook", "Application", "UsedRange", "Debug", "Err",
            "WorksheetFunction"};
    private static final String[] CONSTANTS = {
            "xlLeft", "xlCenter", "xlRight", "xlUp", "xlDown", "xlToLeft", "xlToRight", "vbBlack", "vbRed", "vbGreen", "vbBlue",
            "vbYellow", "vbWhite", "vbMagenta", "vbCyan", "vbCrLf", "vbLf", "vbNewLine", "vbCr", "vbTab", "vbNullString",
            "vbOKOnly", "vbOKCancel", "vbYesNo", "vbYesNoCancel", "vbInformation", "vbExclamation", "vbCritical", "vbQuestion",
            "vbOK", "vbCancel", "vbYes", "vbNo", "vbTrue", "vbFalse", "vbBinaryCompare", "vbTextCompare", "xlGeneral"};

    private static final Map<String, String> KW = new HashMap<>();
    private static final Map<String, String> TY = new HashMap<>();
    private static final Map<String, String> BI = new HashMap<>();
    private static final Map<String, String> OB = new HashMap<>();
    private static final Map<String, String> CO = new HashMap<>();

    static {
        for (String s : KEYWORDS) KW.put(s.toLowerCase(Locale.ROOT), s);
        for (String s : TYPES) TY.put(s.toLowerCase(Locale.ROOT), s);
        for (String s : BUILTINS) BI.put(s.toLowerCase(Locale.ROOT), s);
        for (String s : OBJECTS) OB.put(s.toLowerCase(Locale.ROOT), s);
        for (String s : CONSTANTS) CO.put(s.toLowerCase(Locale.ROOT), s);
    }

    // ------------------------------------------------------------------
    // Tokenising one line
    // ------------------------------------------------------------------

    public static List<Span> spans(String line) {
        List<Span> out = new ArrayList<>();
        int n = line.length();
        int i = 0;
        boolean afterDot = false;
        while (i < n) {
            char c = line.charAt(i);
            if (c == '\'') {
                out.add(new Span(i, n, COMMENT, null));
                return out;
            }
            if (c == '"') {
                int j = i + 1;
                while (j < n) {
                    if (line.charAt(j) == '"') {
                        if (j + 1 < n && line.charAt(j + 1) == '"') {
                            j += 2;
                            continue;
                        }
                        j++;
                        break;
                    }
                    j++;
                }
                out.add(new Span(i, Math.min(j, n), STRING, null));
                i = Math.min(j, n);
                afterDot = false;
            } else if (Character.isDigit(c) || (c == '.' && i + 1 < n && Character.isDigit(line.charAt(i + 1)) && !afterDot)) {
                int j = i;
                if (c == '&' ) j++;
                while (j < n && (Character.isLetterOrDigit(line.charAt(j)) || line.charAt(j) == '.')) j++;
                out.add(new Span(i, j, NUMBER, null));
                i = j;
                afterDot = false;
            } else if (c == '&' && i + 1 < n && (line.charAt(i + 1) == 'H' || line.charAt(i + 1) == 'h')) {
                int j = i + 2;
                while (j < n && Character.digit(line.charAt(j), 16) >= 0) j++;
                out.add(new Span(i, j, NUMBER, null));
                i = j;
                afterDot = false;
            } else if (Character.isLetter(c) || c == '_') {
                int j = i;
                while (j < n && (Character.isLetterOrDigit(line.charAt(j)) || line.charAt(j) == '_')) j++;
                if (j < n && line.charAt(j) == '$') j++;
                String w = line.substring(i, j);
                String low = w.toLowerCase(Locale.ROOT);
                int k = i;
                if (low.equals("rem") && i == firstNonSpace(line)) {
                    out.add(new Span(i, n, COMMENT, null));
                    return out;
                }
                if (afterDot) {
                    out.add(new Span(i, j, MEMBER, null));
                } else if (KW.containsKey(low)) {
                    out.add(new Span(i, j, KEYWORD, KW.get(low)));
                } else if (TY.containsKey(low)) {
                    out.add(new Span(i, j, TYPE, TY.get(low)));
                } else if (CO.containsKey(low)) {
                    out.add(new Span(i, j, CONSTANT, CO.get(low)));
                } else if (OB.containsKey(low)) {
                    out.add(new Span(i, j, OBJECT, null));
                } else if (BI.containsKey(low) || low.equals("date")) {
                    out.add(new Span(i, j, BUILTIN, null));
                } else {
                    out.add(new Span(i, j, PLAIN, null));
                }
                i = j;
                afterDot = false;
            } else {
                afterDot = c == '.';
                if ("+-*/\\^&=<>:".indexOf(c) >= 0) out.add(new Span(i, i + 1, OPERATOR, null));
                i++;
            }
        }
        return out;
    }

    private static int firstNonSpace(String s) {
        int i = 0;
        while (i < s.length() && Character.isWhitespace(s.charAt(i))) i++;
        return i;
    }

    // ------------------------------------------------------------------
    // Indentation
    // ------------------------------------------------------------------

    /** The word(s) that decide a line's role, lower case, after modifiers (Public, Private ...). */
    static String role(String line) {
        List<Span> sp = spans(line);
        int idx = 0;
        while (idx < sp.size() && sp.get(idx).kind == COMMENT) return "";
        if (sp.isEmpty()) return "";
        String first = line.substring(sp.get(0).start, sp.get(0).end).toLowerCase(Locale.ROOT);
        if ((first.equals("public") || first.equals("private") || first.equals("friend") || first.equals("static")) && sp.size() > 1) {
            String second = line.substring(sp.get(1).start, sp.get(1).end).toLowerCase(Locale.ROOT);
            if (second.equals("sub") || second.equals("function")) return second;
            return "";
        }
        if (first.equals("end") && sp.size() > 1) {
            return "end " + line.substring(sp.get(1).start, sp.get(1).end).toLowerCase(Locale.ROOT);
        }
        if (first.equals("select")) return "select";
        if (first.equals("if")) {
            // block If: Then is the last meaningful token
            Span last = null;
            for (Span s : sp) if (s.kind != COMMENT) last = s;
            if (last != null && line.substring(last.start, last.end).equalsIgnoreCase("then")) return "if";
            return "";
        }
        if (first.equals("for") || first.equals("do") || first.equals("while") || first.equals("with")) return first;
        if (first.equals("sub") || first.equals("function")) return first;
        if (first.equals("else") || first.equals("elseif") || first.equals("case") || first.equals("next") || first.equals("loop")
                || first.equals("wend")) return first;
        return "";
    }

    /** Indent level of every line (0 = none). */
    public static int[] levels(String[] lines) {
        int[] out = new int[lines.length];
        int level = 0;
        List<Boolean> selectInCase = new ArrayList<>();
        for (int i = 0; i < lines.length; i++) {
            String r = role(lines[i]);
            int printAt;
            switch (r) {
                case "end sub": case "end function": case "end if": case "end with": case "next": case "loop": case "wend":
                    level = Math.max(0, level - 1);
                    printAt = level;
                    break;
                case "end select":
                    if (!selectInCase.isEmpty()) {
                        if (selectInCase.remove(selectInCase.size() - 1)) level = Math.max(0, level - 1);
                    }
                    level = Math.max(0, level - 1);
                    printAt = level;
                    break;
                case "else": case "elseif":
                    printAt = Math.max(0, level - 1);
                    break;
                case "case": {
                    if (!selectInCase.isEmpty()) {
                        int top = selectInCase.size() - 1;
                        if (selectInCase.get(top)) level = Math.max(0, level - 1);
                        selectInCase.set(top, Boolean.TRUE);
                    }
                    printAt = level;
                    level++;
                    break;
                }
                default:
                    printAt = level;
            }
            out[i] = printAt;
            switch (r) {
                case "sub": case "function": case "if": case "for": case "do": case "while": case "with":
                    level++;
                    break;
                case "select":
                    selectInCase.add(Boolean.FALSE);
                    level++;
                    break;
                default:
            }
        }
        return out;
    }

    /** The closing line for an opener role, or null. */
    public static String closerFor(String role) {
        switch (role) {
            case "sub": return "End Sub";
            case "function": return "End Function";
            case "if": return "End If";
            case "for": return "Next";
            case "do": return "Loop";
            case "while": return "Wend";
            case "select": return "End Select";
            case "with": return "End With";
            default: return null;
        }
    }

    /** True when the block opened at {@code lines[idx]} already has its closing line. */
    public static boolean isClosed(String[] lines, int idx) {
        String open = role(lines[idx]);
        String closer = closerFor(open);
        if (closer == null) return true;
        int depth = 1;
        for (int i = idx + 1; i < lines.length; i++) {
            String r = role(lines[i]);
            if (r.equals(open)) depth++;
            String cr = closer.toLowerCase(Locale.ROOT);
            boolean isCloser = r.equals(cr) || (cr.equals("next") && r.equals("next")) || (cr.equals("loop") && r.equals("loop"))
                    || (cr.equals("wend") && r.equals("wend"));
            if (isCloser) {
                depth--;
                if (depth == 0) return true;
            }
            if ((open.equals("sub") || open.equals("function")) && (r.equals("sub") || r.equals("function"))) return false;
            if (!open.equals("sub") && !open.equals("function") && (r.equals("end sub") || r.equals("end function"))) return false;
        }
        return false;
    }

    /** Re-indents a whole module and fixes keyword capitalisation. */
    public static String format(String source, String unit) {
        String[] lines = source.split("\\r?\\n", -1);
        int[] lv = levels(lines);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < lines.length; i++) {
            String t = fixCase(lines[i].trim());
            if (i > 0) sb.append('\n');
            if (!t.isEmpty()) {
                for (int k = 0; k < lv[i]; k++) sb.append(unit);
                sb.append(t);
            }
        }
        return sb.toString();
    }

    /** Capitalises keywords, types and constants (outside strings and comments). */
    public static String fixCase(String line) {
        StringBuilder sb = new StringBuilder(line);
        for (Span s : spans(line)) {
            if (s.canonical != null && (s.kind == KEYWORD || s.kind == TYPE || s.kind == CONSTANT)) {
                sb.replace(s.start, s.end, s.canonical);
            }
        }
        return sb.toString();
    }

    // ------------------------------------------------------------------
    // Completion data
    // ------------------------------------------------------------------

    /** One entry of the completion list. */
    public static final class Completion {
        public final String text;
        public final String kind; // K keyword, F function, T type, O object, C constant, M member, V variable, P procedure
        public final String detail;

        public Completion(String text, String kind, String detail) {
            this.text = text;
            this.kind = kind;
            this.detail = detail;
        }

        @Override
        public String toString() {
            return text;
        }
    }

    private static final Map<String, String> FUNCTION_HELP = new HashMap<>();

    static {
        Object[][] h = {
                {"Abs", "Abs(number)"}, {"Array", "Array(a, b, ...)"}, {"Asc", "Asc(text)"}, {"CBool", "CBool(x)"}, {"CDbl", "CDbl(x)"},
                {"CInt", "CInt(x)"}, {"CLng", "CLng(x)"}, {"CStr", "CStr(x)"}, {"Chr", "Chr(code)"}, {"DateAdd", "DateAdd(unit, n, date)"},
                {"DateDiff", "DateDiff(unit, d1, d2)"}, {"Format", "Format(value, pattern)"}, {"InStr", "InStr([start,] text, find)"},
                {"InputBox", "InputBox(prompt)"}, {"Int", "Int(number)"}, {"IsNumeric", "IsNumeric(x)"}, {"Join", "Join(array, delimiter)"},
                {"LBound", "LBound(array)"}, {"LCase", "LCase(text)"}, {"Left", "Left(text, n)"}, {"Len", "Len(text)"}, {"MsgBox", "MsgBox message"},
                {"Mid", "Mid(text, start [, length])"}, {"RGB", "RGB(red, green, blue)"}, {"Replace", "Replace(text, find, replacement)"},
                {"Right", "Right(text, n)"}, {"Rnd", "Rnd"}, {"Round", "Round(number [, digits])"}, {"Split", "Split(text, delimiter)"},
                {"Sqr", "Sqr(number)"}, {"Trim", "Trim(text)"}, {"UBound", "UBound(array)"}, {"UCase", "UCase(text)"}, {"Val", "Val(text)"},
                {"Evaluate", "Evaluate(\"formula\")"}, {"Weekday", "Weekday(date)"}, {"Year", "Year(date)"}, {"Month", "Month(date)"},
                {"Day", "Day(date)"}, {"StrReverse", "StrReverse(text)"}, {"Space", "Space(n)"}, {"Timer", "Timer"}};
        for (Object[] e : h) FUNCTION_HELP.put((String) e[0], (String) e[1]);
    }

    private static final String[] RANGE_MEMBERS = {
            "Value", "Formula", "Text", "Address", "Row", "Column", "Count", "Rows", "Columns", "Cells", "Item", "Range", "Offset(",
            "Resize(", "End(", "EntireRow", "EntireColumn", "Font", "Interior", "HorizontalAlignment", "NumberFormat", "ColumnWidth",
            "ClearContents", "Clear", "Select", "Activate", "AutoFit"};
    private static final String[] FONT_MEMBERS = {"Bold", "Italic", "Size", "Color", "Name"};
    private static final String[] INTERIOR_MEMBERS = {"Color"};
    private static final String[] SHEET_MEMBERS = {"Name", "Range(", "Cells(", "UsedRange", "Rows", "Columns", "Activate"};
    private static final String[] APP_MEMBERS = {"ScreenUpdating", "DisplayAlerts", "Calculation", "WorksheetFunction", "ActiveSheet",
            "Worksheets", "Sheets", "ActiveCell", "Selection", "Range(", "Cells("};
    private static final String[] WF_MEMBERS = {"Sum(", "Average(", "Count(", "CountA(", "Max(", "Min(", "Median(", "Product(",
            "CountIf(", "SumIf(", "Round("};
    private static final String[] ERR_MEMBERS = {"Number", "Description", "Clear"};
    private static final String[] DEBUG_MEMBERS = {"Print"};

    /** Members offered after "." for the thing before the dot. */
    public static List<Completion> membersFor(String before) {
        String b = before.toLowerCase(Locale.ROOT);
        String[] src;
        switch (b) {
            case "font": src = FONT_MEMBERS; break;
            case "interior": src = INTERIOR_MEMBERS; break;
            case "application": case "thisworkbook": case "activeworkbook": src = APP_MEMBERS; break;
            case "worksheetfunction": src = WF_MEMBERS; break;
            case "err": src = ERR_MEMBERS; break;
            case "debug": src = DEBUG_MEMBERS; break;
            case "activesheet": case "worksheets": case "sheets": src = SHEET_MEMBERS; break;
            case "range": case "cells": case "activecell": case "selection": case "offset": case "resize": case "end":
            case "usedrange": case "rows": case "columns": case "entirerow": case "entirecolumn": case "item":
                src = RANGE_MEMBERS;
                break;
            default: {
                List<String> all = new ArrayList<>(Arrays.asList(RANGE_MEMBERS));
                all.addAll(Arrays.asList(SHEET_MEMBERS));
                src = all.toArray(new String[0]);
            }
        }
        List<Completion> out = new ArrayList<>();
        Set<String> seen = new HashSet<>();
        for (String m : src) {
            String name = m.endsWith("(") ? m.substring(0, m.length() - 1) : m;
            if (seen.add(name)) out.add(new Completion(name, "M", m.endsWith("(") ? "method" : "property"));
        }
        return out;
    }

    /** Words offered while typing a normal identifier. */
    public static List<Completion> globalCompletions() {
        List<Completion> out = new ArrayList<>();
        for (String s : KEYWORDS) out.add(new Completion(s, "K", "keyword"));
        for (String s : TYPES) out.add(new Completion(s, "T", "type"));
        for (String s : BUILTINS) out.add(new Completion(s, "F", FUNCTION_HELP.getOrDefault(s, "function")));
        for (String s : OBJECTS) out.add(new Completion(s, "O", "object"));
        for (String s : CONSTANTS) out.add(new Completion(s, "C", "constant"));
        out.add(new Completion("Date", "F", "today's date"));
        return out;
    }
}
