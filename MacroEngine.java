import java.awt.Color;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Random;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.SwingConstants;

/**
 * An interpreter for a practical subset of VBA, used by the Macro Studio.
 *
 * <pre>
 *  Procedures   Sub / Function (Public, Private), parameters (ByVal, ByRef, Optional), Call, Exit Sub/Function
 *  Variables    Dim x As Long, Dim a(1 To 10), Dim m(5, 5), ReDim [Preserve], Const, Set
 *  Control      If/ElseIf/Else/End If (and one-line If), For/Next [Step], For Each, Do/Loop While|Until,
 *               While/Wend, Select Case (values, lists, To, Is), With, Exit For/Do, On Error Resume Next
 *  Objects      Range, Cells, ActiveCell, Selection, ActiveSheet, Worksheets("x"), Rows, Columns, UsedRange,
 *               Offset, Resize, End(xlUp), Font, Interior, NumberFormat, Application.WorksheetFunction.*
 *  Functions    text, maths, conversion, date, Split/Join/Array/UBound, Format, InputBox, MsgBox, RGB ...
 * </pre>
 */
public class MacroEngine {

    /** Error raised for any syntax or runtime problem in a macro. */
    public static class MacroException extends Exception {
        public final int line;
        public final int module;

        public MacroException(String message, int line) {
            this(message, line, 0);
        }

        public MacroException(String message, int line, int module) {
            super(line > 0 ? "Line " + line + ": " + message : message);
            this.line = line;
            this.module = module;
        }
    }

    /** Access to the rest of the workbook (optional; without it only the given sheet exists). */
    public interface Workbook {
        List<String> sheetNames();

        SpreadsheetModel sheet(String name);

        String activeSheetName();

        /** Selected block {row0, col0, row1, col1}, or null. */
        int[] selection();
    }

    /** One Sub or Function found by {@link #procedures}. */
    public static final class ProcInfo {
        public final String name;
        public final boolean function;
        public final int line;
        public final int params;
        public final boolean isPrivate;

        ProcInfo(String name, boolean function, int line, int params, boolean isPrivate) {
            this.name = name;
            this.function = function;
            this.line = line;
            this.params = params;
            this.isPrivate = isPrivate;
        }
    }

    private static final int CI = Pattern.CASE_INSENSITIVE;
    private static final Pattern PROC_HEAD = Pattern.compile(
            "^\\s*(?:(public|private|friend)\\s+)?(?:static\\s+)?(sub|function)\\s+([A-Za-z_][A-Za-z0-9_]*)\\s*(\\(([^)]*)\\))?", CI);
    private static final Pattern A1 = Pattern.compile("^\\$?([A-Za-z]{1,3})\\$?(\\d+)$");
    private static final Pattern LEADING_NUMBER = Pattern.compile("^\\s*[-+]?(\\d+\\.?\\d*|\\.\\d+)([eE][-+]?\\d+)?");
    private static final int MAX_DEPTH = 250;
    private static final Map<String, Double> CONSTANTS = new HashMap<>();

    static {
        Object[][] c = {
                {"xlleft", -4131}, {"xlcenter", -4108}, {"xlright", -4152}, {"xlup", -4162}, {"xldown", -4121},
                {"xltoleft", -4159}, {"xltoright", -4161}, {"xlgeneral", 1}, {"xlcalculationmanual", -4135},
                {"xlcalculationautomatic", -4105},
                {"vbblack", 0}, {"vbred", 255}, {"vbgreen", 65280}, {"vbblue", 16711680}, {"vbyellow", 65535},
                {"vbwhite", 16777215}, {"vbmagenta", 16711935}, {"vbcyan", 16776960},
                {"vbokonly", 0}, {"vbokcancel", 1}, {"vbyesno", 4}, {"vbyesnocancel", 3}, {"vbinformation", 64},
                {"vbexclamation", 48}, {"vbcritical", 16}, {"vbquestion", 32},
                {"vbok", 1}, {"vbcancel", 2}, {"vbyes", 6}, {"vbno", 7}, {"vbtrue", -1}, {"vbfalse", 0},
                {"vbbinarycompare", 0}, {"vbtextcompare", 1},
        };
        for (Object[] e : c) CONSTANTS.put((String) e[0], ((Integer) e[1]).doubleValue());
    }

    // ------------------------------------------------------------------
    // Values
    // ------------------------------------------------------------------

    /** A VBA array (1 or 2 dimensions) with its lower bounds. */
    static final class VArray {
        int[] lo;
        int[] hi;
        Object[] data;

        VArray(int[] lo, int[] hi) {
            this.lo = lo;
            this.hi = hi;
            long n = 1;
            for (int i = 0; i < lo.length; i++) n *= (hi[i] - lo[i] + 1L);
            if (n < 0 || n > 5_000_000) throw new IllegalArgumentException("Array too large");
            data = new Object[(int) n];
            Arrays.fill(data, null);
        }

        int offset(int[] idx) {
            int off = 0;
            for (int d = 0; d < lo.length; d++) {
                if (idx[d] < lo[d] || idx[d] > hi[d]) return -1;
                off = off * (hi[d] - lo[d] + 1) + (idx[d] - lo[d]);
            }
            return off;
        }
    }

    /** A block of cells on a sheet. */
    static final class RangeObj {
        final SpreadsheetModel m;
        final String sheetName;
        final int r0, c0, r1, c1;

        RangeObj(SpreadsheetModel m, String sheetName, int r0, int c0, int r1, int c1) {
            this.m = m;
            this.sheetName = sheetName;
            this.r0 = Math.min(r0, r1);
            this.c0 = Math.min(c0, c1);
            this.r1 = Math.max(r0, r1);
            this.c1 = Math.max(c0, c1);
        }

        boolean single() {
            return r0 == r1 && c0 == c1;
        }

        long count() {
            return (long) (r1 - r0 + 1) * (c1 - c0 + 1);
        }
    }

    static final class SheetObj {
        final SpreadsheetModel m;
        final String name;

        SheetObj(SpreadsheetModel m, String name) {
            this.m = m;
            this.name = name;
        }
    }

    /** Range.Rows / Range.Columns (for .Count and indexing). */
    static final class DimObj {
        final RangeObj r;
        final boolean rows;

        DimObj(RangeObj r, boolean rows) {
            this.r = r;
            this.rows = rows;
        }
    }

    static final class FontObj {
        final RangeObj r;

        FontObj(RangeObj r) {
            this.r = r;
        }
    }

    static final class InteriorObj {
        final RangeObj r;

        InteriorObj(RangeObj r) {
            this.r = r;
        }
    }

    private static final class AppObj {
    }

    private static final class WfObj {
    }

    private static final class ErrObj {
    }

    /** Names that are never treated as undeclared variables. */
    private static final java.util.Set<String> KNOWN = new java.util.HashSet<>(Arrays.asList(
            "range", "cells", "rows", "columns", "activecell", "selection", "activesheet", "thisworkbook", "activeworkbook",
            "application", "worksheets", "sheets", "usedrange", "err", "msgbox", "inputbox", "evaluate", "vbcrlf", "vblf", "vbnewline",
            "vbcr", "vbtab", "vbnullstring", "pi", "rnd", "now", "date", "time", "timer", "worksheetfunction",
            "int", "fix", "abs", "sgn", "sqr", "sqrt", "exp", "log", "sin", "cos", "tan", "atn", "round", "len", "ucase", "lcase", "trim", "ltrim",
            "rtrim", "left", "right", "mid", "instr", "instrrev", "replace", "space", "string", "chr", "asc", "strreverse", "strcomp", "str",
            "cstr", "val", "cint", "clng", "cdbl", "csng", "ccur", "cdate", "cbool", "isnumeric", "isempty", "isnull", "isarray", "isdate",
            "iserror", "isobject", "typename", "year", "month", "day", "weekday", "dateadd", "datediff", "dateserial", "rgb", "format",
            "formatnumber", "formatcurrency", "formatpercent", "split", "join", "array", "ubound", "lbound", "max", "min"));

    private static boolean knownName(String low) {
        return KNOWN.contains(low) || CONSTANTS.containsKey(low);
    }

    // ------------------------------------------------------------------
    // Tokens and lines
    // ------------------------------------------------------------------

    private static final class Tok {
        final char type; // n number, s string, i identifier, o operator
        final String text;

        Tok(char type, String text) {
            this.type = type;
            this.text = text;
        }

        boolean isOp(String o) {
            return type == 'o' && text.equals(o);
        }

        boolean isWord(String w) {
            return type == 'i' && text.equalsIgnoreCase(w);
        }
    }

    private static final class Line {
        final List<Tok> toks;
        final int no;
        final int module;

        Line(List<Tok> toks, int no, int module) {
            this.toks = toks;
            this.no = no;
            this.module = module;
        }

        String first() {
            return toks.isEmpty() ? "" : (toks.get(0).type == 'i' ? toks.get(0).text.toLowerCase(Locale.ROOT) : "");
        }

        String second() {
            return toks.size() < 2 || toks.get(1).type != 'i' ? "" : toks.get(1).text.toLowerCase(Locale.ROOT);
        }
    }

    // ------------------------------------------------------------------
    // Program structure
    // ------------------------------------------------------------------

    private interface Ex {
        Object ev(Frame f) throws MacroException;
    }

    private interface St {
        void run(Frame f) throws MacroException;
    }

    private static final class S {
        final int line;
        final int module;
        final St st;

        S(int line, int module, St st) {
            this.line = line;
            this.module = module;
            this.st = st;
        }
    }

    private static final class Param {
        String name;
        boolean byVal;
        boolean optional;
        Ex dflt;
    }

    private static final class Proc {
        String name;
        boolean function;
        boolean isPrivate;
        List<Param> params = new ArrayList<>();
        S[] body;
        int line;
        int module;
    }

    private static final class Frame {
        final Map<String, Object> vars = new HashMap<>();
        final Proc proc;
        boolean resumeNext;
        final List<Object> withStack = new ArrayList<>();

        Frame(Proc p) {
            this.proc = p;
        }
    }

    private static final class ExitSig extends RuntimeException {
        final String kind;

        ExitSig(String kind) {
            super(null, null, false, false);
            this.kind = kind;
        }
    }

    private static final class EndSig extends RuntimeException {
        EndSig() {
            super(null, null, false, false);
        }
    }

    // ------------------------------------------------------------------
    // Engine state
    // ------------------------------------------------------------------

    private final SpreadsheetModel model;
    private final Consumer<String> out;
    private final Random random = new Random();
    private final Map<String, Proc> procs = new HashMap<>();
    private final List<S> moduleInit = new ArrayList<>();
    private final Frame globals = new Frame(null);
    private final Map<String, Object> consts = new HashMap<>();
    private final ErrObj errObj = new ErrObj();
    private long errNumber;
    private String errDesc = "";
    private int curLine;
    private int curModule;
    private long steps;
    private int depth;
    private long deadline;

    private Consumer<String> msgBox;
    private Function<String, String> inputBox;
    private Workbook workbook;
    private BooleanSupplier cancelled = () -> false;
    private long maxMillis = 20_000;
    private long maxSteps = 5_000_000;
    private String activeName = "Sheet1";

    public MacroEngine(SpreadsheetModel model, Consumer<String> out) {
        this.model = model;
        this.out = out;
    }

    /** Where MsgBox goes (default: the output text "MsgBox: ..."). */
    public void setMsgBoxHandler(Consumer<String> h) {
        msgBox = h;
    }

    /** Where InputBox asks (return null when cancelled; default: an empty string). */
    public void setInputBoxHandler(Function<String, String> h) {
        inputBox = h;
    }

    public void setWorkbook(Workbook w) {
        workbook = w;
        if (w != null && w.activeSheetName() != null) activeName = w.activeSheetName();
    }

    public void setCancelCheck(BooleanSupplier c) {
        cancelled = c;
    }

    public void setLimits(long maxMillis, long maxSteps) {
        this.maxMillis = maxMillis;
        this.maxSteps = maxSteps;
    }

    private MacroException err(String message) {
        return new MacroException(message, curLine, curModule);
    }

    // ------------------------------------------------------------------
    // Quick procedure scan (for the editor's drop-down; never throws)
    // ------------------------------------------------------------------

    public static List<ProcInfo> procedures(String source) {
        List<ProcInfo> out = new ArrayList<>();
        String[] lines = source.split("\\r?\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            String l = lines[i];
            if (l.trim().startsWith("'")) continue;
            Matcher m = PROC_HEAD.matcher(l);
            if (m.find()) {
                String ps = m.group(5);
                int n = ps == null || ps.trim().isEmpty() ? 0 : ps.split(",").length;
                out.add(new ProcInfo(m.group(3), m.group(2).equalsIgnoreCase("function"), i + 1, n,
                        m.group(1) != null && m.group(1).equalsIgnoreCase("private")));
            }
        }
        return out;
    }

    // ------------------------------------------------------------------
    // Running
    // ------------------------------------------------------------------

    /** Runs the first Sub in {@code source} (Subs without parameters first). */
    public int run(String source) throws MacroException {
        load(java.util.Collections.singletonList(source));
        String best = null;
        for (ProcInfo p : procedures(source)) {
            if (p.function) continue;
            if (p.params == 0) {
                best = p.name;
                break;
            }
            if (best == null) best = p.name;
        }
        if (best == null) throw new MacroException("There is no Sub to run", 0);
        return runProcedure(best);
    }

    /** Parses all modules (replacing any earlier program). */
    public void load(List<String> modules) throws MacroException {
        procs.clear();
        initDone = false;
        moduleInit.clear();
        globals.vars.clear();
        consts.clear();
        for (int mi = 0; mi < modules.size(); mi++) {
            curModule = mi;
            List<Line> lines = lex(modules.get(mi), mi);
            new Parser(lines, mi).parseModule();
        }
    }

    public boolean hasProcedure(String name) {
        return procs.containsKey(name.toLowerCase(Locale.ROOT));
    }

    /** Runs a procedure by name with no arguments. @return executed statements */
    public int runProcedure(String name) throws MacroException {
        if (model.isProtected()) {
            throw new MacroException("This sheet is protected. Unprotect it before running a macro.", 0);
        }
        Proc p = procs.get(name.toLowerCase(Locale.ROOT));
        if (p == null) throw new MacroException("Sub or Function not defined: " + name, 0);
        steps = 0;
        depth = 0;
        deadline = System.currentTimeMillis() + maxMillis;
        model.beginBatch();
        try {
            runInit();
            callProc(p, new ArrayList<>(), null, null);
        } catch (EndSig e) {
            // End statement: stop quietly
        } catch (ExitSig e) {
            // stray Exit: ignore
        } finally {
            model.endBatch();
        }
        return (int) Math.min(Integer.MAX_VALUE, steps);
    }

    private boolean initDone;

    private void runInit() throws MacroException {
        if (initDone) return;
        initDone = true;
        for (S s : moduleInit) execOne(s, globals);
    }

    /**
     * Immediate window: "? expr" prints a value, anything else runs as a statement
     * (procedures from the loaded modules can be called).
     */
    public void immediate(String text) throws MacroException {
        String t = text.trim();
        if (t.isEmpty()) return;
        curLine = 0;
        curModule = 0;
        steps = 0;
        depth = 0;
        deadline = System.currentTimeMillis() + maxMillis;
        boolean print = t.startsWith("?");
        if (print) t = "Debug.Print " + t.substring(1);
        List<Line> lines = lex(t, -1);
        Parser ps = new Parser(lines, -1);
        S[] body = ps.parseStatements();
        model.beginBatch();
        try {
            runInit();
            Frame f = new Frame(null);
            execBlock(body, f);
        } catch (EndSig e) {
            // ignore
        } finally {
            model.endBatch();
        }
    }

    // ------------------------------------------------------------------
    // Lexer
    // ------------------------------------------------------------------

    private List<Line> lex(String source, int module) throws MacroException {
        String[] raw = source.split("\\r?\\n", -1);
        List<Line> out = new ArrayList<>();
        for (int i = 0; i < raw.length; i++) {
            int startNo = i + 1;
            String text = raw[i];
            // line continuation: " _" at the end joins the next line
            while (stripComment(text).stripTrailing().endsWith(" _") && i + 1 < raw.length) {
                String cur = stripComment(text).stripTrailing();
                text = cur.substring(0, cur.length() - 1) + " " + raw[++i].trim();
            }
            curLine = startNo;
            curModule = Math.max(0, module);
            List<Tok> toks = tokenize(stripComment(text));
            if (toks.isEmpty()) continue;
            splitStatements(toks, startNo, module, out);
        }
        return out;
    }

    /** Splits a physical line at ':' (except inside a one-line If, which keeps the whole line). */
    private void splitStatements(List<Tok> toks, int no, int module, List<Line> out) {
        boolean oneLineIf = toks.get(0).isWord("if");
        boolean caseLine = toks.get(0).isWord("case");
        if (oneLineIf || caseLine) {
            out.add(new Line(toks, no, module));
            return;
        }
        List<Tok> cur = new ArrayList<>();
        for (Tok t : toks) {
            if (t.isOp(":")) {
                if (!cur.isEmpty()) out.add(new Line(cur, no, module));
                cur = new ArrayList<>();
            } else {
                cur.add(t);
            }
        }
        if (!cur.isEmpty()) out.add(new Line(cur, no, module));
    }

    private static String stripComment(String s) {
        boolean inString = false;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == '"') inString = !inString;
            else if (c == '\'' && !inString) return s.substring(0, i);
        }
        String t = s.trim().toLowerCase(Locale.ROOT);
        if (t.equals("rem") || t.startsWith("rem ")) return "";
        return s;
    }

    private List<Tok> tokenize(String s) throws MacroException {
        List<Tok> toks = new ArrayList<>();
        int i = 0;
        int len = s.length();
        while (i < len) {
            char c = s.charAt(i);
            if (Character.isWhitespace(c)) {
                i++;
            } else if (c == '&' && i + 2 < len && (s.charAt(i + 1) == 'H' || s.charAt(i + 1) == 'h')
                    && Character.digit(s.charAt(i + 2), 16) >= 0) {
                int j = i + 2;
                while (j < len && Character.digit(s.charAt(j), 16) >= 0) j++;
                toks.add(new Tok('n', String.valueOf((double) Long.parseLong(s.substring(i + 2, j), 16))));
                i = j;
                if (i < len && s.charAt(i) == '&') i++;
            } else if (Character.isDigit(c) || (c == '.' && i + 1 < len && Character.isDigit(s.charAt(i + 1)))) {
                int j = i;
                while (j < len && (Character.isDigit(s.charAt(j)) || s.charAt(j) == '.')) j++;
                if (j < len && (s.charAt(j) == 'e' || s.charAt(j) == 'E') && j + 1 < len
                        && (Character.isDigit(s.charAt(j + 1)) || ((s.charAt(j + 1) == '-' || s.charAt(j + 1) == '+')
                        && j + 2 < len && Character.isDigit(s.charAt(j + 2))))) {
                    j += 2;
                    while (j < len && Character.isDigit(s.charAt(j))) j++;
                }
                toks.add(new Tok('n', s.substring(i, j)));
                i = j;
                if (i < len && (s.charAt(i) == '#' || s.charAt(i) == '%' || s.charAt(i) == '!')) i++;
            } else if (c == '"') {
                StringBuilder sb = new StringBuilder();
                i++;
                while (true) {
                    if (i >= len) throw err("Unterminated string literal");
                    char d = s.charAt(i);
                    if (d == '"') {
                        if (i + 1 < len && s.charAt(i + 1) == '"') {
                            sb.append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    sb.append(d);
                    i++;
                }
                toks.add(new Tok('s', sb.toString()));
            } else if (Character.isLetter(c) || c == '_') {
                int j = i;
                while (j < len && (Character.isLetterOrDigit(s.charAt(j)) || s.charAt(j) == '_')) j++;
                toks.add(new Tok('i', s.substring(i, j)));
                i = j;
                if (i < len && s.charAt(i) == '$') i++; // Left$, Mid$ ...
            } else if (c == '[') { // [A1] shorthand
                int j = s.indexOf(']', i);
                if (j < 0) throw err("Missing ]");
                toks.add(new Tok('i', "Range"));
                toks.add(new Tok('o', "("));
                toks.add(new Tok('s', s.substring(i + 1, j)));
                toks.add(new Tok('o', ")"));
                i = j + 1;
            } else if (s.startsWith(":=", i)) {
                toks.add(new Tok('o', ":="));
                i += 2;
            } else if (i + 1 < len && (s.startsWith("<>", i) || s.startsWith("<=", i) || s.startsWith(">=", i))) {
                toks.add(new Tok('o', s.substring(i, i + 2)));
                i += 2;
            } else if ("+-*/\\^&=<>(),.:;".indexOf(c) >= 0) {
                toks.add(new Tok('o', String.valueOf(c)));
                i++;
            } else {
                throw err("Unexpected character '" + c + "'");
            }
        }
        return toks;
    }

    // ------------------------------------------------------------------
    // Parser
    // ------------------------------------------------------------------

    private final class Parser {
        final List<Line> lines;
        final int module;
        int p;

        Parser(List<Line> lines, int module) {
            this.lines = lines;
            this.module = module;
        }

        Line cur() {
            return p < lines.size() ? lines.get(p) : null;
        }

        MacroException at(Line l, String msg) {
            return new MacroException(msg, l == null ? curLine : l.no, Math.max(0, module));
        }

        void parseModule() throws MacroException {
            while (p < lines.size()) {
                Line l = lines.get(p);
                String k = l.first();
                int kk = k.equals("public") || k.equals("private") || k.equals("friend") || k.equals("static") ? 1 : 0;
                String w = kk == 1 ? l.second() : k;
                if (w.equals("sub") || w.equals("function")) {
                    parseProc();
                } else if (w.equals("option") || w.equals("attribute") || w.equals("declare") || w.equals("type") || w.equals("enum")) {
                    p++;
                } else {
                    S s = parseStatement();
                    if (s != null) moduleInit.add(s);
                }
            }
        }

        void parseProc() throws MacroException {
            Line head = lines.get(p);
            int i = 0;
            Proc pr = new Proc();
            pr.module = Math.max(0, module);
            pr.line = head.no;
            Tok t = head.toks.get(i);
            if (t.isWord("public") || t.isWord("private") || t.isWord("friend")) {
                pr.isPrivate = t.isWord("private");
                t = head.toks.get(++i);
            }
            if (t.isWord("static")) t = head.toks.get(++i);
            pr.function = t.isWord("function");
            i++;
            if (i >= head.toks.size() || head.toks.get(i).type != 'i') throw at(head, "Procedure name expected");
            pr.name = head.toks.get(i++).text;
            if (i < head.toks.size() && head.toks.get(i).isOp("(")) {
                i++;
                TokCur tc = new TokCur(head.toks, i);
                while (!tc.atEnd() && !tc.peek().isOp(")")) {
                    Param pa = new Param();
                    if (tc.acceptWord("optional")) pa.optional = true;
                    if (tc.acceptWord("byval")) pa.byVal = true;
                    else tc.acceptWord("byref");
                    if (tc.acceptWord("paramarray")) pa.optional = true;
                    if (tc.atEnd() || tc.peek().type != 'i') throw at(head, "Parameter name expected");
                    pa.name = tc.next().text.toLowerCase(Locale.ROOT);
                    if (tc.acceptOp("(")) tc.expectOp(")");
                    if (tc.acceptWord("as")) {
                        tc.next();
                        if (tc.acceptOp(".")) tc.next();
                    }
                    if (tc.acceptOp("=")) pa.dflt = new ExprParser(tc, head).parseExpr();
                    pr.params.add(pa);
                    if (!tc.acceptOp(",")) break;
                }
            }
            p++;
            S[] body = parseBlock(new String[]{"end sub", "end function"});
            Line end = cur();
            if (end == null) throw at(head, (pr.function ? "Function" : "Sub") + " without End " + (pr.function ? "Function" : "Sub"));
            p++;
            pr.body = body;
            if (procs.containsKey(pr.name.toLowerCase(Locale.ROOT))) throw at(head, "Ambiguous name: " + pr.name);
            procs.put(pr.name.toLowerCase(Locale.ROOT), pr);
        }

        String key(Line l) {
            String k = l.first();
            if (k.equals("end")) return "end " + l.second();
            if (k.equals("exit")) return "exit";
            return k;
        }

        S[] parseStatements() throws MacroException {
            List<S> out = new ArrayList<>();
            while (p < lines.size()) {
                S s = parseStatement();
                if (s != null) out.add(s);
            }
            return out.toArray(new S[0]);
        }

        /** Parses lines until one whose key is in {@code stops} (not consumed). */
        S[] parseBlock(String[] stops) throws MacroException {
            List<S> out = new ArrayList<>();
            while (p < lines.size()) {
                Line l = lines.get(p);
                String k = key(l);
                for (String s : stops) if (k.equals(s)) return out.toArray(new S[0]);
                if (k.equals("end sub") || k.equals("end function")) return out.toArray(new S[0]); // block left open
                S s = parseStatement();
                if (s != null) out.add(s);
            }
            return out.toArray(new S[0]);
        }

        S wrap(Line l, St st) {
            return new S(l.no, Math.max(0, module), st);
        }

        S parseStatement() throws MacroException {
            Line l = lines.get(p);
            String k = l.first();
            switch (k) {
                case "dim": case "static": case "private": case "public": case "const": return parseDim(l);
                case "redim": return parseRedim(l);
                case "if": return parseIf(l);
                case "for": return l.second().equals("each") ? parseForEach(l) : parseFor(l);
                case "do": return parseDo(l);
                case "while": return parseWhile(l);
                case "select": return parseSelect(l);
                case "with": return parseWith(l);
                case "exit": return parseExit(l);
                case "on": return parseOnError(l);
                case "end":
                    if (l.toks.size() == 1) {
                        p++;
                        return wrap(l, f -> {
                            throw new EndSig();
                        });
                    }
                    throw at(l, "Unexpected End " + l.second());
                case "next": throw at(l, "Next without For");
                case "loop": throw at(l, "Loop without Do");
                case "wend": throw at(l, "Wend without While");
                case "else": case "elseif": throw at(l, "Else without If");
                case "case": throw at(l, "Case without Select Case");
                case "set": {
                    p++;
                    ExprParser ep = new ExprParser(new TokCur(l.toks, 1), l);
                    return assignOrCall(l, ep);
                }
                case "call": {
                    p++;
                    ExprParser ep = new ExprParser(new TokCur(l.toks, 1), l);
                    Ex e = ep.parseExpr();
                    ep.expectEnd();
                    return wrap(l, f -> {
                        e.ev(f);
                    });
                }
                case "stop": case "randomize": case "option": case "attribute": {
                    p++;
                    return null;
                }
                case "goto": case "gosub": case "resume": case "return":
                    throw at(l, "GoTo / Resume labels are not supported - use loops and If instead");
                default:
                    p++;
                    return parseSimple(l, 0);
            }
        }

        /** Assignment, procedure call or object call on one line (from token index {@code from}). */
        S parseSimple(Line l, int from) throws MacroException {
            ExprParser ep = new ExprParser(new TokCur(l.toks, from), l);
            return assignOrCall(l, ep);
        }

        S assignOrCall(Line l, ExprParser ep) throws MacroException {
            TokCur tc = ep.tc;
            Tok first = tc.peek();
            if (first == null) throw at(l, "Statement expected");
            if (first.type != 'i' && !first.isOp(".")) throw at(l, "Syntax error: cannot understand '" + l.toks.get(0).text + "'");
            String lowerFirst = first.text.toLowerCase(Locale.ROOT);
            if (lowerFirst.equals("msgbox")) {
                tc.next();
                return msgBoxStmt(l, tc, ep);
            }
            if (lowerFirst.equals("debug") && tc.peekAt(1) != null && tc.peekAt(1).isOp(".")) {
                tc.next();
                tc.next();
                String what = tc.atEnd() ? "" : tc.next().text.toLowerCase(Locale.ROOT);
                if (!what.equals("print")) throw at(l, "Unsupported Debug member: " + what);
                List<Ex> parts = new ArrayList<>();
                List<String> seps = new ArrayList<>();
                while (!tc.atEnd()) {
                    parts.add(ep.parseExpr());
                    if (tc.acceptOp(";")) seps.add("");
                    else if (tc.acceptOp(",")) seps.add("\t");
                    else break;
                }
                ep.expectEnd();
                return wrap(l, f -> {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < parts.size(); i++) {
                        sb.append(str(parts.get(i).ev(f)));
                        if (i < seps.size() && i < parts.size() - 1) sb.append(seps.get(i));
                    }
                    out.accept(sb.toString());
                });
            }
            // paren-less call: Name arg1, arg2   (identifier followed by something that starts an expression)
            Tok second = tc.peekAt(1);
            boolean call = first.type == 'i' && second != null && !second.isOp("=") && !second.isOp("(") && !second.isOp(".")
                    && !second.isOp(":=") && (second.type == 'i' || second.type == 'n' || second.type == 's' || second.isOp("-"))
                    && !hasEqualsAtTop(l, tc.p);
            if (call) {
                String name = tc.next().text;
                List<Ex> args = new ArrayList<>();
                while (!tc.atEnd()) {
                    args.add(ep.parseArg());
                    if (!tc.acceptOp(",")) break;
                }
                ep.expectEnd();
                Ref r = new Ref(null, name, args, l.no);
                return wrap(l, f -> {
                    r.ev(f);
                });
            }
            Ex target = ep.parsePostfix();
            if (tc.acceptOp("=")) {
                Ex rhs = ep.parseExpr();
                ep.expectEnd();
                if (!(target instanceof Ref)) throw at(l, "Cannot assign to this expression");
                Ref r = (Ref) target;
                return wrap(l, f -> r.assign(f, rhs.ev(f)));
            }
            // paren-less call of a member with arguments:  rng.Select / obj.Method a, b
            List<Ex> extra = new ArrayList<>();
            if (!tc.atEnd()) {
                while (!tc.atEnd()) {
                    extra.add(ep.parseArg());
                    if (!tc.acceptOp(",")) break;
                }
                ep.expectEnd();
            }
            if (!(target instanceof Ref)) throw at(l, "Syntax error");
            Ref r = (Ref) target;
            if (!extra.isEmpty()) r = r.withExtraArgs(extra);
            final Ref rr = r;
            return wrap(l, f -> {
                if (rr.base == null && rr.args == null && !hasVar(f, rr.low) && !procs.containsKey(rr.low) && !knownName(rr.low)) {
                    throw err("Sub or Function not defined: " + rr.name);
                }
                rr.ev(f);
            });
        }

        boolean hasEqualsAtTop(Line l, int from) {
            int depth = 0;
            for (int i = from; i < l.toks.size(); i++) {
                Tok t = l.toks.get(i);
                if (t.isOp("(")) depth++;
                else if (t.isOp(")")) depth--;
                else if (depth == 0 && t.isOp("=")) return true;
            }
            return false;
        }

        S msgBoxStmt(Line l, TokCur tc, ExprParser ep) throws MacroException {
            List<Ex> args = new ArrayList<>();
            boolean paren = tc.acceptOp("(");
            while (!tc.atEnd() && !tc.peek().isOp(")")) {
                args.add(ep.parseArg());
                if (!tc.acceptOp(",")) break;
            }
            if (paren) tc.expectOp(")");
            ep.expectEnd();
            if (args.isEmpty()) throw at(l, "MsgBox needs a message");
            return wrap(l, f -> {
                String text = str(args.get(0).ev(f));
                if (args.size() > 1) args.get(1).ev(f);
                showMsg(text);
            });
        }

        S parseDim(Line l) throws MacroException {
            p++;
            TokCur tc = new TokCur(l.toks, 0);
            boolean isConst = false;
            if (tc.acceptWord("public") || tc.acceptWord("private") || tc.acceptWord("static")) {
                if (tc.peek() != null && tc.peek().isWord("sub") || tc.peek() != null && tc.peek().isWord("function")) {
                    throw at(l, "Procedure declaration not expected here");
                }
            }
            if (tc.acceptWord("dim")) {
                // fine
            } else if (tc.acceptWord("const")) {
                isConst = true;
            }
            if (tc.atEnd()) throw at(l, "Variable name expected");
            ExprParser ep = new ExprParser(tc, l);
            List<St> parts = new ArrayList<>();
            do {
                if (tc.atEnd() || tc.peek().type != 'i') throw at(l, "Variable name expected");
                final String name = tc.next().text.toLowerCase(Locale.ROOT);
                List<Ex[]> dims = null;
                if (tc.acceptOp("(")) {
                    dims = new ArrayList<>();
                    if (!tc.acceptOp(")")) {
                        do {
                            Ex a = ep.parseExpr();
                            Ex b = null;
                            if (tc.acceptWord("to")) {
                                b = ep.parseExpr();
                            }
                            dims.add(new Ex[]{b == null ? null : a, b == null ? a : b});
                        } while (tc.acceptOp(","));
                        tc.expectOp(")");
                    }
                }
                String type = "variant";
                if (tc.acceptWord("as")) {
                    tc.acceptWord("new");
                    type = tc.atEnd() ? "variant" : tc.next().text.toLowerCase(Locale.ROOT);
                    while (tc.acceptOp(".")) tc.next();
                    if (tc.acceptOp("*")) tc.next(); // String * 10
                }
                Ex init = null;
                if (isConst || tc.peek() != null && tc.peek().isOp("=")) {
                    tc.expectOp("=");
                    init = ep.parseExpr();
                }
                final String ftype = type;
                final List<Ex[]> fdims = dims;
                final Ex finit = init;
                final boolean fconst = isConst;
                parts.add(f -> {
                    Object v;
                    if (fdims != null && !fdims.isEmpty()) {
                        v = makeArray(fdims, f);
                    } else if (fdims != null) {
                        v = null; // dynamic array, ReDim later
                    } else if (finit != null) {
                        v = finit.ev(f);
                    } else {
                        v = defaultFor(ftype);
                    }
                    if (fconst) consts.put(name, v);
                    else declare(f, name, v);
                });
            } while (tc.acceptOp(","));
            ep.expectEnd();
            return wrap(l, f -> {
                for (St s : parts) s.run(f);
            });
        }

        S parseRedim(Line l) throws MacroException {
            p++;
            TokCur tc = new TokCur(l.toks, 1);
            boolean preserve = tc.acceptWord("preserve");
            ExprParser ep = new ExprParser(tc, l);
            List<St> parts = new ArrayList<>();
            do {
                final String name = tc.next().text.toLowerCase(Locale.ROOT);
                tc.expectOp("(");
                List<Ex[]> dims = new ArrayList<>();
                do {
                    Ex a = ep.parseExpr();
                    Ex b = null;
                    if (tc.acceptWord("to")) b = ep.parseExpr();
                    dims.add(new Ex[]{b == null ? null : a, b == null ? a : b});
                } while (tc.acceptOp(","));
                tc.expectOp(")");
                if (tc.acceptWord("as")) tc.next();
                parts.add(f -> {
                    VArray na = makeArray(dims, f);
                    if (preserve) {
                        Object old = lookupVar(f, name);
                        if (old instanceof VArray) {
                            VArray oa = (VArray) old;
                            if (oa.lo.length == na.lo.length) {
                                int[] idx = new int[oa.lo.length];
                                copyArray(oa, na, idx, 0);
                            }
                        }
                    }
                    declare(f, name, na);
                });
            } while (tc.acceptOp(","));
            ep.expectEnd();
            return wrap(l, f -> {
                for (St s : parts) s.run(f);
            });
        }

        S parseIf(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            Ex cond = ep.parseExpr();
            if (!tc.acceptWord("then")) throw at(l, "Expected Then");
            if (!tc.atEnd()) { // one-line If
                p++;
                List<Tok> rest = l.toks.subList(tc.p, l.toks.size());
                int elseAt = -1;
                int depth = 0;
                for (int i = 0; i < rest.size(); i++) {
                    Tok t = rest.get(i);
                    if (t.isOp("(")) depth++;
                    else if (t.isOp(")")) depth--;
                    else if (depth == 0 && t.isWord("else")) {
                        elseAt = i;
                        break;
                    }
                }
                S[] thenS = inline(l, elseAt < 0 ? rest : rest.subList(0, elseAt));
                S[] elseS = elseAt < 0 ? null : inline(l, rest.subList(elseAt + 1, rest.size()));
                return wrap(l, f -> {
                    if (toBool(cond.ev(f))) execBlock(thenS, f);
                    else if (elseS != null) execBlock(elseS, f);
                });
            }
            p++;
            List<Ex> conds = new ArrayList<>();
            List<S[]> blocks = new ArrayList<>();
            conds.add(cond);
            blocks.add(parseBlock(new String[]{"elseif", "else", "end if"}));
            S[] elseBlock = null;
            while (true) {
                Line n = cur();
                if (n == null) throw at(l, "If without End If");
                String k = key(n);
                if (!k.equals("elseif") && !k.equals("else") && !k.equals("end if")) throw at(l, "If without End If");
                if (k.equals("elseif")) {
                    TokCur t2 = new TokCur(n.toks, 1);
                    ExprParser e2 = new ExprParser(t2, n);
                    Ex c2 = e2.parseExpr();
                    if (!t2.acceptWord("then")) throw at(n, "Expected Then");
                    p++;
                    conds.add(c2);
                    blocks.add(parseBlock(new String[]{"elseif", "else", "end if"}));
                } else if (k.equals("else")) {
                    p++;
                    elseBlock = parseBlock(new String[]{"end if"});
                } else {
                    p++;
                    break;
                }
            }
            final S[] eb = elseBlock;
            return wrap(l, f -> {
                for (int i = 0; i < conds.size(); i++) {
                    if (toBool(conds.get(i).ev(f))) {
                        execBlock(blocks.get(i), f);
                        return;
                    }
                }
                if (eb != null) execBlock(eb, f);
            });
        }

        /** Statements of a one-line If branch (separated by ':'). */
        S[] inline(Line l, List<Tok> toks) throws MacroException {
            List<Line> sub = new ArrayList<>();
            List<Tok> curr = new ArrayList<>();
            for (Tok t : toks) {
                if (t.isOp(":")) {
                    if (!curr.isEmpty()) sub.add(new Line(curr, l.no, l.module));
                    curr = new ArrayList<>();
                } else {
                    curr.add(t);
                }
            }
            if (!curr.isEmpty()) sub.add(new Line(curr, l.no, l.module));
            if (sub.isEmpty()) throw at(l, "Statement expected after Then/Else");
            Parser ps = new Parser(sub, module);
            return ps.parseStatements();
        }

        S parseFor(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            if (tc.atEnd() || tc.peek().type != 'i') throw at(l, "Syntax error in For statement");
            final String var = tc.next().text.toLowerCase(Locale.ROOT);
            if (!tc.acceptOp("=")) throw at(l, "Syntax error in For statement");
            Ex start = ep.parseExpr();
            if (!tc.acceptWord("to")) throw at(l, "Syntax error in For statement");
            Ex end = ep.parseExpr();
            Ex step = null;
            if (tc.acceptWord("step")) step = ep.parseExpr();
            ep.expectEnd();
            p++;
            S[] body = parseBlock(new String[]{"next"});
            Line n = cur();
            if (n == null || !key(n).equals("next")) throw at(l, "For without Next");
            p++;
            final Ex fstep = step;
            return wrap(l, f -> {
                double a = toNum(start.ev(f));
                double b = toNum(end.ev(f));
                double st = fstep == null ? 1 : toNum(fstep.ev(f));
                if (st == 0) throw err("Step cannot be zero");
                for (double v = a; st > 0 ? v <= b : v >= b; v += st) {
                    setVar(f, var, v);
                    try {
                        execBlock(body, f);
                    } catch (ExitSig e) {
                        if (!e.kind.equals("for")) throw e;
                        break;
                    }
                    Object cur = lookupVar(f, var);
                    if (cur instanceof Double) v = (Double) cur; // body may change the counter
                }
            });
        }

        S parseForEach(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 2);
            ExprParser ep = new ExprParser(tc, l);
            if (tc.atEnd() || tc.peek().type != 'i') throw at(l, "Syntax error in For Each");
            final String var = tc.next().text.toLowerCase(Locale.ROOT);
            if (!tc.acceptWord("in")) throw at(l, "Expected In");
            Ex coll = ep.parseExpr();
            ep.expectEnd();
            p++;
            S[] body = parseBlock(new String[]{"next"});
            if (cur() == null || !key(cur()).equals("next")) throw at(l, "For Each without Next");
            p++;
            return wrap(l, f -> {
                List<Object> items = itemsOf(coll.ev(f));
                for (Object it : items) {
                    setVar(f, var, it);
                    try {
                        execBlock(body, f);
                    } catch (ExitSig e) {
                        if (!e.kind.equals("for")) throw e;
                        break;
                    }
                }
            });
        }

        S parseDo(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            Ex preCond = null;
            boolean preUntil = false;
            if (tc.acceptWord("while")) preCond = ep.parseExpr();
            else if (tc.acceptWord("until")) {
                preCond = ep.parseExpr();
                preUntil = true;
            }
            ep.expectEnd();
            p++;
            S[] body = parseBlock(new String[]{"loop"});
            Line n = cur();
            if (n == null || !key(n).equals("loop")) throw at(l, "Do without Loop");
            p++;
            Ex postCond = null;
            boolean postUntil = false;
            TokCur t2 = new TokCur(n.toks, 1);
            ExprParser e2 = new ExprParser(t2, n);
            if (t2.acceptWord("while")) postCond = e2.parseExpr();
            else if (t2.acceptWord("until")) {
                postCond = e2.parseExpr();
                postUntil = true;
            }
            e2.expectEnd();
            if (preCond != null && postCond != null) throw at(n, "Loop cannot have a condition when Do has one");
            final Ex pc = preCond;
            final boolean pu = preUntil;
            final Ex qc = postCond;
            final boolean qu = postUntil;
            return wrap(l, f -> {
                while (true) {
                    if (pc != null && toBool(pc.ev(f)) == pu) break;
                    try {
                        execBlock(body, f);
                    } catch (ExitSig e) {
                        if (!e.kind.equals("do")) throw e;
                        break;
                    }
                    if (qc != null && toBool(qc.ev(f)) == qu) break;
                    tick();
                }
            });
        }

        S parseWhile(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            Ex cond = ep.parseExpr();
            ep.expectEnd();
            p++;
            S[] body = parseBlock(new String[]{"wend"});
            if (cur() == null || !key(cur()).equals("wend")) throw at(l, "While without Wend");
            p++;
            return wrap(l, f -> {
                while (toBool(cond.ev(f))) {
                    try {
                        execBlock(body, f);
                    } catch (ExitSig e) {
                        if (!e.kind.equals("do")) throw e;
                        break;
                    }
                    tick();
                }
            });
        }

        S parseSelect(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            if (!tc.acceptWord("case")) throw at(l, "Expected Case after Select");
            Ex subject = ep.parseExpr();
            ep.expectEnd();
            p++;
            List<List<Ex[]>> tests = new ArrayList<>(); // each case: list of {kind, a, b}
            List<String[]> kinds = new ArrayList<>();
            List<S[]> bodies = new ArrayList<>();
            S[] elseBody = null;
            // anything before the first Case is ignored (VBA allows only comments)
            while (true) {
                Line n = cur();
                if (n == null || key(n).startsWith("end s") && !key(n).equals("end select")) throw at(l, "Select Case without End Select");
                String k = key(n);
                if (k.equals("end select")) {
                    p++;
                    break;
                }
                if (!k.equals("case")) throw at(n, "Expected Case or End Select");
                p++;
                TokCur ct = new TokCur(n.toks, 1);
                ExprParser ce = new ExprParser(ct, n);
                if (ct.acceptWord("else")) {
                    S[] first = ct.acceptOp(":") ? inline(n, n.toks.subList(ct.p, n.toks.size())) : new S[0];
                    elseBody = concat(first, parseBlock(new String[]{"case", "end select"}));
                    continue;
                }
                List<Ex[]> alts = new ArrayList<>();
                List<String> ks = new ArrayList<>();
                do {
                    if (ct.acceptWord("is")) {
                        String op = ct.atEnd() ? "" : ct.next().text;
                        if (!op.matches("<>|<=|>=|=|<|>")) throw at(n, "Comparison operator expected after Is");
                        alts.add(new Ex[]{ce.parseExpr(), null});
                        ks.add(op);
                    } else {
                        Ex a = ce.parseExpr();
                        if (ct.acceptWord("to")) {
                            alts.add(new Ex[]{a, ce.parseExpr()});
                            ks.add("to");
                        } else {
                            alts.add(new Ex[]{a, null});
                            ks.add("=");
                        }
                    }
                } while (ct.acceptOp(","));
                S[] first = new S[0];
                if (ct.acceptOp(":")) first = inline(n, n.toks.subList(ct.p, n.toks.size()));
                else ce.expectEnd();
                tests.add(alts);
                kinds.add(ks.toArray(new String[0]));
                bodies.add(concat(first, parseBlock(new String[]{"case", "end select"})));
            }
            final S[] eb = elseBody;
            return wrap(l, f -> {
                Object v = subject.ev(f);
                for (int i = 0; i < tests.size(); i++) {
                    List<Ex[]> alts = tests.get(i);
                    String[] ks = kinds.get(i);
                    for (int j = 0; j < alts.size(); j++) {
                        Object a = alts.get(j)[0].ev(f);
                        boolean hit;
                        if (ks[j].equals("to")) {
                            Object b = alts.get(j)[1].ev(f);
                            hit = compare(v, a, ">=") && compare(v, b, "<=");
                        } else {
                            hit = compare(v, a, ks[j]);
                        }
                        if (hit) {
                            execBlock(bodies.get(i), f);
                            return;
                        }
                    }
                }
                if (eb != null) execBlock(eb, f);
            });
        }

        S[] concat(S[] a, S[] b) {
            S[] r = Arrays.copyOf(a, a.length + b.length);
            System.arraycopy(b, 0, r, a.length, b.length);
            return r;
        }

        S parseWith(Line l) throws MacroException {
            TokCur tc = new TokCur(l.toks, 1);
            ExprParser ep = new ExprParser(tc, l);
            Ex obj = ep.parseExpr();
            ep.expectEnd();
            p++;
            S[] body = parseBlock(new String[]{"end with"});
            if (cur() == null || !key(cur()).equals("end with")) throw at(l, "With without End With");
            p++;
            return wrap(l, f -> {
                f.withStack.add(obj.ev(f));
                try {
                    execBlock(body, f);
                } finally {
                    f.withStack.remove(f.withStack.size() - 1);
                }
            });
        }

        S parseExit(Line l) throws MacroException {
            p++;
            String what = l.second();
            if (!what.equals("for") && !what.equals("do") && !what.equals("sub") && !what.equals("function")) {
                throw at(l, "Exit must be followed by For, Do, Sub or Function");
            }
            return wrap(l, f -> {
                throw new ExitSig(what);
            });
        }

        S parseOnError(Line l) throws MacroException {
            p++;
            String rest = l.toks.size() > 2 ? l.toks.get(2).text.toLowerCase(Locale.ROOT) : "";
            if (rest.equals("resume")) return wrap(l, f -> f.resumeNext = true);
            if (rest.equals("goto") && l.toks.size() > 3 && l.toks.get(3).text.equals("0")) return wrap(l, f -> f.resumeNext = false);
            throw at(l, "Only On Error Resume Next and On Error GoTo 0 are supported");
        }
    }

    // ------------------------------------------------------------------
    // Token cursor and expression parser
    // ------------------------------------------------------------------

    private final class TokCur {
        final List<Tok> t;
        int p;

        TokCur(List<Tok> t, int p) {
            this.t = t;
            this.p = p;
        }

        boolean atEnd() {
            return p >= t.size();
        }

        Tok peek() {
            return p < t.size() ? t.get(p) : null;
        }

        Tok peekAt(int off) {
            return p + off < t.size() ? t.get(p + off) : null;
        }

        Tok next() {
            return t.get(p++);
        }

        boolean acceptOp(String op) {
            if (peek() != null && peek().isOp(op)) {
                p++;
                return true;
            }
            return false;
        }

        boolean acceptWord(String w) {
            if (peek() != null && peek().isWord(w)) {
                p++;
                return true;
            }
            return false;
        }

        void expectOp(String op) throws MacroException {
            if (!acceptOp(op)) throw err("Expected '" + op + "'");
        }
    }

    private final class ExprParser {
        final TokCur tc;
        final Line line;

        ExprParser(TokCur tc, Line line) {
            this.tc = tc;
            this.line = line;
        }

        MacroException e(String msg) {
            return new MacroException(msg, line.no, Math.max(0, line.module));
        }

        void expectEnd() throws MacroException {
            if (!tc.atEnd()) throw e("Unexpected '" + tc.peek().text + "'");
        }

        /** An argument; named arguments (Name:=value) are accepted and the name is ignored. */
        Ex parseArg() throws MacroException {
            if (tc.peek() != null && tc.peek().type == 'i' && tc.peekAt(1) != null && tc.peekAt(1).isOp(":=")) {
                tc.p += 2;
            }
            return parseExpr();
        }

        Ex parseExpr() throws MacroException {
            return orExpr();
        }

        private Ex orExpr() throws MacroException {
            Ex left = andExpr();
            while (true) {
                if (tc.acceptWord("or")) {
                    Ex l = left;
                    Ex r = andExpr();
                    left = f -> logic(l.ev(f), r.ev(f), 0);
                } else if (tc.acceptWord("xor")) {
                    Ex l = left;
                    Ex r = andExpr();
                    left = f -> logic(l.ev(f), r.ev(f), 2);
                } else {
                    return left;
                }
            }
        }

        private Ex andExpr() throws MacroException {
            Ex left = notExpr();
            while (tc.acceptWord("and")) {
                Ex l = left;
                Ex r = notExpr();
                left = f -> logic(l.ev(f), r.ev(f), 1);
            }
            return left;
        }

        private Ex notExpr() throws MacroException {
            if (tc.acceptWord("not")) {
                Ex x = notExpr();
                return f -> !toBool(x.ev(f));
            }
            return cmpExpr();
        }

        private Ex cmpExpr() throws MacroException {
            Ex left = concat();
            for (String op : new String[]{"<>", "<=", ">=", "=", "<", ">"}) {
                if (tc.peek() != null && tc.peek().isOp(op)) {
                    tc.p++;
                    Ex l = left;
                    Ex r = concat();
                    return f -> compare(l.ev(f), r.ev(f), op);
                }
            }
            if (tc.peek() != null && tc.peek().isWord("like")) {
                tc.p++;
                Ex l = left;
                Ex r = concat();
                return f -> like(str(l.ev(f)), str(r.ev(f)));
            }
            return left;
        }

        private Ex concat() throws MacroException {
            Ex left = addSub();
            while (tc.acceptOp("&")) {
                Ex l = left;
                Ex r = addSub();
                left = f -> str(l.ev(f)) + str(r.ev(f));
            }
            return left;
        }

        private Ex addSub() throws MacroException {
            Ex left = mulDiv();
            while (true) {
                if (tc.acceptOp("+")) {
                    Ex l = left;
                    Ex r = mulDiv();
                    left = f -> {
                        Object a = l.ev(f);
                        Object b = r.ev(f);
                        if (a instanceof String && b instanceof String) return (String) a + b;
                        return toNum(a) + toNum(b);
                    };
                } else if (tc.acceptOp("-")) {
                    Ex l = left;
                    Ex r = mulDiv();
                    left = f -> toNum(l.ev(f)) - toNum(r.ev(f));
                } else {
                    return left;
                }
            }
        }

        private Ex mulDiv() throws MacroException {
            Ex left = unary();
            while (true) {
                if (tc.acceptOp("*")) {
                    Ex l = left;
                    Ex r = unary();
                    left = f -> toNum(l.ev(f)) * toNum(r.ev(f));
                } else if (tc.acceptOp("/")) {
                    Ex l = left;
                    Ex r = unary();
                    left = f -> {
                        double a = toNum(l.ev(f));
                        double d = toNum(r.ev(f));
                        if (d == 0) throw err("Division by zero");
                        return a / d;
                    };
                } else if (tc.acceptOp("\\")) {
                    Ex l = left;
                    Ex r = unary();
                    left = f -> {
                        long a = Math.round(toNum(l.ev(f)));
                        long d = Math.round(toNum(r.ev(f)));
                        if (d == 0) throw err("Division by zero");
                        return (double) (a / d);
                    };
                } else if (tc.acceptWord("mod")) {
                    Ex l = left;
                    Ex r = unary();
                    left = f -> {
                        long a = Math.round(toNum(l.ev(f)));
                        long d = Math.round(toNum(r.ev(f)));
                        if (d == 0) throw err("Division by zero");
                        return (double) (a % d);
                    };
                } else {
                    return left;
                }
            }
        }

        private Ex unary() throws MacroException {
            if (tc.acceptOp("-")) {
                Ex x = unary();
                return f -> -toNum(x.ev(f));
            }
            if (tc.acceptOp("+")) return unary();
            return power();
        }

        private Ex power() throws MacroException {
            Ex base = parsePostfix();
            if (tc.acceptOp("^")) {
                Ex ex = unary();
                return f -> Math.pow(toNum(base.ev(f)), toNum(ex.ev(f)));
            }
            return base;
        }

        /** A primary followed by any number of .Member / (args) parts. */
        Ex parsePostfix() throws MacroException {
            Tok tok = tc.peek();
            if (tok == null) throw e("Unexpected end of expression");
            Ex cur;
            if (tok.type == 'n') {
                tc.p++;
                final double d;
                try {
                    d = Double.parseDouble(tok.text);
                } catch (NumberFormatException ex) {
                    throw e("Invalid number '" + tok.text + "'");
                }
                cur = f -> d;
            } else if (tok.type == 's') {
                tc.p++;
                final String s = tok.text;
                cur = f -> s;
            } else if (tc.acceptOp("(")) {
                Ex inner = parseExpr();
                tc.expectOp(")");
                cur = inner;
                // a parenthesised expression can still be followed by .Member
            } else if (tok.isOp(".")) {
                tc.p++;
                Tok nm = tc.atEnd() ? null : tc.next();
                if (nm == null || nm.type != 'i') throw e("Member name expected after '.'");
                List<Ex> args = parseArgsOpt();
                cur = new Ref(new WithEx(), nm.text, args, line.no);
            } else if (tok.type == 'i') {
                tc.p++;
                String name = tok.text;
                String low = name.toLowerCase(Locale.ROOT);
                if (low.equals("true")) cur = f -> Boolean.TRUE;
                else if (low.equals("false")) cur = f -> Boolean.FALSE;
                else if (low.equals("empty") || low.equals("nothing") || low.equals("null")) cur = f -> null;
                else {
                    List<Ex> args = parseArgsOpt();
                    cur = new Ref(null, name, args, line.no);
                }
            } else {
                throw e("Unexpected '" + tok.text + "'");
            }
            while (tc.peek() != null && tc.peek().isOp(".")) {
                tc.p++;
                Tok nm = tc.atEnd() ? null : tc.next();
                if (nm == null || nm.type != 'i') throw e("Member name expected after '.'");
                List<Ex> args = parseArgsOpt();
                cur = new Ref(cur, nm.text, args, line.no);
            }
            return cur;
        }

        /** "(a, b)" directly after a name, or null when there are no parentheses. */
        List<Ex> parseArgsOpt() throws MacroException {
            if (tc.peek() == null || !tc.peek().isOp("(")) return null;
            tc.p++;
            List<Ex> args = new ArrayList<>();
            if (tc.acceptOp(")")) return args;
            do {
                if (tc.peek() != null && tc.peek().isOp(",")) { // omitted argument
                    args.add(f -> null);
                    continue;
                }
                args.add(parseArg());
            } while (tc.acceptOp(","));
            tc.expectOp(")");
            return args;
        }
    }

    private static final class WithEx implements Ex {
        @Override
        public Object ev(Frame f) throws MacroException {
            if (f.withStack.isEmpty()) throw new MacroException("Invalid use of . outside a With block", 0);
            return f.withStack.get(f.withStack.size() - 1);
        }
    }

    // ------------------------------------------------------------------
    // References: variables, arrays, procedure calls, built-ins and object members
    // ------------------------------------------------------------------

    private final class Ref implements Ex {
        final Ex base;
        final String name;
        final String low;
        final List<Ex> args;
        final int line;

        Ref(Ex base, String name, List<Ex> args, int line) {
            this.base = base;
            this.name = name;
            this.low = name.toLowerCase(Locale.ROOT);
            this.args = args;
            this.line = line;
        }

        Ref withExtraArgs(List<Ex> extra) {
            List<Ex> all = new ArrayList<>();
            if (args != null) all.addAll(args);
            all.addAll(extra);
            return new Ref(base, name, all, line);
        }

        List<Object> evalArgs(Frame f) throws MacroException {
            List<Object> out = new ArrayList<>();
            if (args != null) for (Ex a : args) out.add(a.ev(f));
            return out;
        }

        @Override
        public Object ev(Frame f) throws MacroException {
            if (base != null) {
                Object obj = base.ev(f);
                return member(obj, low, evalArgs(f), false, null);
            }
            // local / global variable (or array element)
            Object var = lookupVar(f, low);
            boolean isVar = var != null || hasVar(f, low);
            if (isVar) {
                if (args == null) return var;
                if (var == null && procs.containsKey(low)) return callProc(procs.get(low), evalArgs(f), args, f);
                if (var instanceof VArray) return arrayGet((VArray) var, evalArgs(f));
                if (var instanceof RangeObj || var instanceof SheetObj) return member(var, "cells", evalArgs(f), false, null);
                throw err("Subscript out of range");
            }
            if (consts.containsKey(low)) return consts.get(low);
            Proc pr = procs.get(low);
            if (pr != null) return callProc(pr, evalArgs(f), args, f);
            if (args == null && !knownName(low)) return null; // undeclared variable reads as Empty
            return root(low, evalArgs(f), args != null, f);
        }

        void assign(Frame f, Object value) throws MacroException {
            if (base == null) {
                if (args == null) {
                    if (consts.containsKey(low)) throw err("Cannot assign to constant " + name);
                    setVar(f, low, value);
                    return;
                }
                Object var = lookupVar(f, low);
                if (var instanceof VArray) {
                    arraySet((VArray) var, evalArgs(f), value);
                    return;
                }
                Object target = ev(f);
                if (target instanceof RangeObj) {
                    member(target, "value", new ArrayList<>(), true, value);
                    return;
                }
                throw err("Cannot assign to " + name);
            }
            Object obj = base.ev(f);
            if (args != null && !args.isEmpty()) {
                // obj.Prop(args) = value  e.g. Range("A1").Cells(1,2) = 5  or arr.Item(i)
                Object target = member(obj, low, evalArgs(f), false, null);
                if (target instanceof RangeObj) {
                    member(target, "value", new ArrayList<>(), true, value);
                    return;
                }
                throw err("Cannot assign to " + name);
            }
            member(obj, low, new ArrayList<>(), true, value);
        }
    }

    // ------------------------------------------------------------------
    // Variables
    // ------------------------------------------------------------------

    private static Object lookupVar(Frame f, String name, Map<String, Object> globalVars) {
        Object v = f.vars.get(name);
        if (v != null || f.vars.containsKey(name)) return v;
        return globalVars.get(name);
    }

    private Object lookupVar(Frame f, String name) {
        return lookupVar(f, name, globals.vars);
    }

    private boolean hasVar(Frame f, String name) {
        return f.vars.containsKey(name) || globals.vars.containsKey(name);
    }

    private void declare(Frame f, String name, Object v) {
        if (f.proc == null) globals.vars.put(name, v);
        else f.vars.put(name, v);
    }

    private void setVar(Frame f, String name, Object v) {
        if (f.vars.containsKey(name) || !globals.vars.containsKey(name)) {
            if (f.proc == null) globals.vars.put(name, v);
            else f.vars.put(name, v);
        } else {
            globals.vars.put(name, v);
        }
    }

    private Object defaultFor(String type) {
        switch (type) {
            case "string": return "";
            case "boolean": return Boolean.FALSE;
            case "integer": case "long": case "double": case "single": case "byte": case "currency": case "longlong": case "date":
                return 0.0;
            default: return null;
        }
    }

    private VArray makeArray(List<Ex[]> dims, Frame f) throws MacroException {
        int n = dims.size();
        if (n > 2) throw err("Only one- and two-dimensional arrays are supported");
        int[] lo = new int[n];
        int[] hi = new int[n];
        for (int i = 0; i < n; i++) {
            lo[i] = dims.get(i)[0] == null ? 0 : (int) toNum(dims.get(i)[0].ev(f));
            hi[i] = (int) toNum(dims.get(i)[1].ev(f));
            if (hi[i] < lo[i] - 1) throw err("Subscript out of range");
        }
        try {
            return new VArray(lo, hi);
        } catch (IllegalArgumentException ex) {
            throw err("Out of memory: " + ex.getMessage());
        }
    }

    private void copyArray(VArray from, VArray to, int[] idx, int d) {
        if (d == from.lo.length) {
            int a = from.offset(idx);
            int b = to.offset(idx);
            if (a >= 0 && b >= 0) to.data[b] = from.data[a];
            return;
        }
        for (int i = from.lo[d]; i <= from.hi[d]; i++) {
            idx[d] = i;
            copyArray(from, to, idx, d + 1);
        }
    }

    private Object arrayGet(VArray a, List<Object> idx) throws MacroException {
        int[] ix = new int[idx.size()];
        for (int i = 0; i < ix.length; i++) ix[i] = (int) Math.round(toNum(idx.get(i)));
        if (ix.length != a.lo.length) throw err("Wrong number of array subscripts");
        int off = a.offset(ix);
        if (off < 0) throw err("Subscript out of range");
        return a.data[off];
    }

    private void arraySet(VArray a, List<Object> idx, Object v) throws MacroException {
        int[] ix = new int[idx.size()];
        for (int i = 0; i < ix.length; i++) ix[i] = (int) Math.round(toNum(idx.get(i)));
        if (ix.length != a.lo.length) throw err("Wrong number of array subscripts");
        int off = a.offset(ix);
        if (off < 0) throw err("Subscript out of range");
        a.data[off] = v;
    }

    // ------------------------------------------------------------------
    // Executing
    // ------------------------------------------------------------------

    private void tick() throws MacroException {
        if (++steps > maxSteps) throw err("Execution limit exceeded (infinite loop?)");
        if ((steps & 1023) == 0) {
            if (cancelled.getAsBoolean()) throw err("Macro stopped");
            if (System.currentTimeMillis() > deadline) throw err("Macro took too long and was stopped (over " + maxMillis / 1000 + " seconds)");
        }
    }

    private void execBlock(S[] body, Frame f) throws MacroException {
        for (S s : body) execOne(s, f);
    }

    private void execOne(S s, Frame f) throws MacroException {
        curLine = s.line;
        curModule = s.module;
        tick();
        try {
            s.st.run(f);
        } catch (MacroException ex) {
            if (f.resumeNext) {
                errNumber = 1004;
                errDesc = ex.getMessage();
                return;
            }
            throw ex;
        } catch (ArithmeticException | IndexOutOfBoundsException | ClassCastException | IllegalArgumentException ex) {
            MacroException me = new MacroException("Run-time error: " + ex.getMessage(), s.line, s.module);
            if (f.resumeNext) return;
            throw me;
        } catch (StackOverflowError ex) {
            throw new MacroException("Out of stack space", s.line, s.module);
        }
    }

    private Object callProc(Proc p, List<Object> args, List<Ex> argExprs, Frame caller) throws MacroException {
        if (++depth > MAX_DEPTH) {
            depth = 0;
            throw err("Out of stack space (too many nested calls)");
        }
        int savedLine = curLine;
        int savedModule = curModule;
        Frame f = new Frame(p);
        if (args.size() > p.params.size()) throw err("Wrong number of arguments for " + p.name);
        for (int i = 0; i < p.params.size(); i++) {
            Param pa = p.params.get(i);
            Object v;
            if (i < args.size() && !(args.get(i) == null && pa.optional && pa.dflt != null)) {
                v = args.get(i);
            } else if (pa.optional) {
                v = pa.dflt == null ? null : pa.dflt.ev(f);
            } else {
                throw err("Argument not optional: " + pa.name + " in " + p.name);
            }
            f.vars.put(pa.name, v);
        }
        if (p.function) f.vars.put(p.name.toLowerCase(Locale.ROOT), null);
        try {
            execBlock(p.body, f);
        } catch (ExitSig e) {
            if (!e.kind.equals("sub") && !e.kind.equals("function")) throw e;
        } finally {
            depth--;
        }
        // ByRef write-back for plain variable arguments
        if (argExprs != null && caller != null) {
            for (int i = 0; i < p.params.size() && i < argExprs.size(); i++) {
                Param pa = p.params.get(i);
                Ex ae = argExprs.get(i);
                if (!pa.byVal && ae instanceof Ref) {
                    Ref r = (Ref) ae;
                    if (r.base == null && r.args == null && (caller.vars.containsKey(r.low) || globals.vars.containsKey(r.low))) {
                        setVar(caller, r.low, f.vars.get(pa.name));
                    }
                }
            }
        }
        curLine = savedLine;
        curModule = savedModule;
        return p.function ? f.vars.get(p.name.toLowerCase(Locale.ROOT)) : null;
    }

    private void showMsg(String text) {
        if (msgBox != null) msgBox.accept(text);
        else out.accept("MsgBox: " + text);
    }

    // ------------------------------------------------------------------
    // Value helpers
    // ------------------------------------------------------------------

    private double toNum(Object v) throws MacroException {
        if (v == null) return 0;
        if (v instanceof Double) return (Double) v;
        if (v instanceof Boolean) return ((Boolean) v) ? -1 : 0;
        if (v instanceof RangeObj) return toNum(rangeValue((RangeObj) v));
        String s = v.toString().trim();
        if (s.isEmpty()) return 0;
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            throw err("Type mismatch: '" + s + "' is not a number");
        }
    }

    private boolean toBool(Object v) throws MacroException {
        if (v instanceof Boolean) return (Boolean) v;
        if (v instanceof String) {
            String s = ((String) v).trim();
            if (s.equalsIgnoreCase("true")) return true;
            if (s.equalsIgnoreCase("false")) return false;
        }
        return toNum(v) != 0;
    }

    private String str(Object v) throws MacroException {
        if (v == null) return "";
        if (v instanceof Double) return FormulaEngine.formatNumber((Double) v);
        if (v instanceof Boolean) return ((Boolean) v) ? "True" : "False";
        if (v instanceof RangeObj) return str(rangeValue((RangeObj) v));
        if (v instanceof SheetObj) return ((SheetObj) v).name;
        if (v instanceof VArray) throw err("Type mismatch: an array cannot be used as text");
        return v.toString();
    }

    private static String inputString(Object v) {
        if (v == null) return "";
        if (v instanceof Boolean) return ((Boolean) v) ? "TRUE" : "FALSE";
        if (v instanceof Double) {
            double d = (Double) v;
            if (d == Math.rint(d) && Math.abs(d) < 1e15) return String.valueOf((long) d);
            return FormulaEngine.formatNumber(d);
        }
        return v.toString();
    }

    private Object logic(Object a, Object b, int op) throws MacroException {
        if (a instanceof Boolean && b instanceof Boolean || (!(a instanceof Double) && !(b instanceof Double))) {
            boolean x = toBool(a);
            boolean y = toBool(b);
            return op == 0 ? x || y : op == 1 ? x && y : x ^ y;
        }
        long x = (long) toNum(a);
        long y = (long) toNum(b);
        return (double) (op == 0 ? x | y : op == 1 ? x & y : x ^ y);
    }

    private boolean compare(Object a, Object b, String op) throws MacroException {
        if (a instanceof RangeObj) a = rangeValue((RangeObj) a);
        if (b instanceof RangeObj) b = rangeValue((RangeObj) b);
        int cmp;
        if (a instanceof String && b instanceof String) cmp = ((String) a).compareToIgnoreCase((String) b);
        else if ((a instanceof String && !isNumeric((String) a)) || (b instanceof String && !isNumeric((String) b))) {
            cmp = str(a).compareToIgnoreCase(str(b));
        } else cmp = Double.compare(toNum(a), toNum(b));
        switch (op) {
            case "=": return cmp == 0;
            case "<>": return cmp != 0;
            case "<": return cmp < 0;
            case ">": return cmp > 0;
            case "<=": return cmp <= 0;
            default: return cmp >= 0;
        }
    }

    private static boolean isNumeric(String s) {
        String t = s.trim();
        if (t.isEmpty()) return true; // Empty compares as 0 / ""
        try {
            Double.parseDouble(t);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private boolean like(String s, String pat) {
        StringBuilder re = new StringBuilder();
        for (int i = 0; i < pat.length(); i++) {
            char c = pat.charAt(i);
            if (c == '*') re.append(".*");
            else if (c == '?') re.append('.');
            else if (c == '#') re.append("\\d");
            else if (c == '[') {
                int j = pat.indexOf(']', i);
                if (j < 0) re.append("\\[");
                else {
                    String cls = pat.substring(i + 1, j);
                    if (cls.startsWith("!")) cls = "^" + cls.substring(1);
                    re.append('[').append(cls).append(']');
                    i = j;
                }
            } else re.append(Pattern.quote(String.valueOf(c)));
        }
        return Pattern.compile(re.toString(), CI | Pattern.DOTALL).matcher(s).matches();
    }

    private List<Object> itemsOf(Object c) throws MacroException {
        List<Object> items = new ArrayList<>();
        if (c instanceof VArray) {
            items.addAll(Arrays.asList(((VArray) c).data));
        } else if (c instanceof RangeObj) {
            RangeObj r = (RangeObj) c;
            if (r.count() > 2_000_000) throw err("Range too large for For Each");
            for (int row = r.r0; row <= r.r1; row++) {
                for (int col = r.c0; col <= r.c1; col++) items.add(new RangeObj(r.m, r.sheetName, row, col, row, col));
            }
        } else if (c instanceof DimObj) {
            DimObj d = (DimObj) c;
            int n = d.rows ? d.r.r1 - d.r.r0 + 1 : d.r.c1 - d.r.c0 + 1;
            for (int i = 0; i < n; i++) items.add(d.rows
                    ? new RangeObj(d.r.m, d.r.sheetName, d.r.r0 + i, d.r.c0, d.r.r0 + i, d.r.c1)
                    : new RangeObj(d.r.m, d.r.sheetName, d.r.r0, d.r.c0 + i, d.r.r1, d.r.c0 + i));
        } else {
            throw err("For Each needs a range or an array");
        }
        return items;
    }

    // ------------------------------------------------------------------
    // Ranges
    // ------------------------------------------------------------------

    private Object rangeValue(RangeObj r) throws MacroException {
        if (r.single()) {
            Object v = r.m.getComputedValue(r.r0, r.c0);
            return v == null ? "" : v;
        }
        if (r.count() > 1_000_000) throw err("Range too large to read into an array");
        VArray a = new VArray(new int[]{1, 1}, new int[]{r.r1 - r.r0 + 1, r.c1 - r.c0 + 1});
        for (int i = r.r0; i <= r.r1; i++) {
            for (int j = r.c0; j <= r.c1; j++) {
                Object v = r.m.getComputedValue(i, j);
                a.data[(i - r.r0) * (r.c1 - r.c0 + 1) + (j - r.c0)] = v == null ? "" : v;
            }
        }
        return a;
    }

    private void writeRange(RangeObj r, Object v, boolean formula) throws MacroException {
        if (v instanceof VArray) {
            VArray a = (VArray) v;
            int rows = a.hi[0] - a.lo[0] + 1;
            int cols = a.lo.length == 2 ? a.hi[1] - a.lo[1] + 1 : 1;
            boolean oneD = a.lo.length == 1;
            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < cols; j++) {
                    int rr = r.r0 + (oneD ? 0 : i);
                    int cc = r.c0 + (oneD ? i : j);
                    if (!r.m.inBounds(rr, cc)) throw err("Subscript out of range");
                    int off = oneD ? i : i * cols + j;
                    r.m.setInput(rr, cc, inputString(a.data[off]));
                }
            }
            return;
        }
        String text = inputString(v);
        for (int i = r.r0; i <= r.r1; i++) for (int j = r.c0; j <= r.c1; j++) r.m.setInput(i, j, text);
    }

    private void styleEach(RangeObj r, Consumer<CellStyle> change) {
        for (int i = r.r0; i <= r.r1; i++) for (int j = r.c0; j <= r.c1; j++) r.m.updateStyle(i, j, change);
    }

    private static Color fromVba(double value) {
        int n = (int) value;
        return new Color(n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF);
    }

    private static double toVba(Color c) {
        return c == null ? 16777215 : (c.getRed() | (c.getGreen() << 8) | (c.getBlue() << 16));
    }

    private RangeObj cellRef(String a1, SpreadsheetModel m, String sheet) throws MacroException {
        Matcher mm = A1.matcher(a1.trim());
        if (!mm.matches()) throw err("Invalid cell reference '" + a1 + "'");
        int r = Integer.parseInt(mm.group(2)) - 1;
        int c = FormulaEngine.parseColumn(mm.group(1));
        return checked(new RangeObj(m, sheet, r, c, r, c));
    }

    private RangeObj checked(RangeObj r) throws MacroException {
        if (!r.m.inBounds(r.r0, r.c0) || !r.m.inBounds(r.r1, r.c1)) {
            throw err("Subscript out of range: " + FormulaEngine.cellRef(Math.max(r.r0, 0), Math.max(r.c0, 0)));
        }
        return r;
    }

    private RangeObj parseRange(List<Object> a, SpreadsheetModel m, String sheet) throws MacroException {
        if (a.isEmpty() || a.size() > 2) throw err("Range needs one or two arguments");
        if (a.size() == 1 && a.get(0) instanceof RangeObj) return (RangeObj) a.get(0);
        if (a.size() == 2 && a.get(0) instanceof RangeObj && a.get(1) instanceof RangeObj) {
            RangeObj x = (RangeObj) a.get(0);
            RangeObj y = (RangeObj) a.get(1);
            return checked(new RangeObj(x.m, x.sheetName, Math.min(x.r0, y.r0), Math.min(x.c0, y.c0), Math.max(x.r1, y.r1), Math.max(x.c1, y.c1)));
        }
        String s = str(a.get(0)).trim();
        String second = a.size() == 2 ? str(a.get(1)).trim() : null;
        if (second == null && s.contains(":")) {
            String[] parts = s.split(":");
            if (parts.length != 2) throw err("Invalid range '" + s + "'");
            s = parts[0];
            second = parts[1];
        }
        // whole column / row references like "A:A" or "3:3"
        if (second != null && s.matches("[A-Za-z]{1,3}") && second.matches("[A-Za-z]{1,3}")) {
            return checked(new RangeObj(m, sheet, 0, FormulaEngine.parseColumn(s), m.getRowCount() - 1, FormulaEngine.parseColumn(second)));
        }
        if (second != null && s.matches("\\d+") && second.matches("\\d+")) {
            return checked(new RangeObj(m, sheet, Integer.parseInt(s) - 1, 0, Integer.parseInt(second) - 1, m.getColumnCount() - 1));
        }
        RangeObj first = cellRef(s, m, sheet);
        if (second == null) return first;
        RangeObj sec = cellRef(second, m, sheet);
        return checked(new RangeObj(m, sheet, first.r0, first.c0, sec.r0, sec.c0));
    }

    private RangeObj cellsOf(RangeObj base, SpreadsheetModel m, String sheet, List<Object> a) throws MacroException {
        if (a.isEmpty()) {
            return base != null ? base : checked(new RangeObj(m, sheet, 0, 0, m.getRowCount() - 1, m.getColumnCount() - 1));
        }
        int rOff = base == null ? 0 : base.r0;
        int cOff = base == null ? 0 : base.c0;
        int ri;
        int ci;
        if (a.size() == 1) { // single index: counts across the range row by row
            int idx = (int) Math.round(toNum(a.get(0))) - 1;
            int w = base == null ? m.getColumnCount() : base.c1 - base.c0 + 1;
            ri = idx / w;
            ci = idx % w;
        } else {
            ri = (int) Math.round(toNum(a.get(0))) - 1;
            Object co = a.get(1);
            ci = (co instanceof String && !((String) co).matches("\\d+")) ? FormulaEngine.parseColumn((String) co)
                    : (int) Math.round(toNum(co)) - 1;
        }
        RangeObj r = new RangeObj(m, sheet, rOff + ri, cOff + ci, rOff + ri, cOff + ci);
        if (base != null && (ri < 0 || ci < 0 || rOff + ri > base.r1 || cOff + ci > base.c1)) {
            // Excel lets Cells() reach outside a range; keep that, only the sheet bounds matter
        }
        return checked(r);
    }

    private boolean emptyCell(SpreadsheetModel m, int r, int c) {
        return m.getInput(r, c).isEmpty();
    }

    private RangeObj endOf(RangeObj r, int dir) throws MacroException {
        int row = r.r0;
        int col = r.c0;
        int dr = dir == -4162 ? -1 : dir == -4121 ? 1 : 0;
        int dc = dir == -4159 ? -1 : dir == -4161 ? 1 : 0;
        if (dr == 0 && dc == 0) throw err("Unsupported End direction");
        SpreadsheetModel m = r.m;
        int nr = row + dr;
        int nc = col + dc;
        if (!m.inBounds(nr, nc)) return new RangeObj(m, r.sheetName, row, col, row, col);
        if (!emptyCell(m, row, col) && !emptyCell(m, nr, nc)) {
            // run to the end of the filled block
            while (m.inBounds(row + dr, col + dc) && !emptyCell(m, row + dr, col + dc)) {
                row += dr;
                col += dc;
            }
        } else {
            // jump to the next filled cell (or the sheet edge)
            row = nr;
            col = nc;
            while (m.inBounds(row + dr, col + dc) && emptyCell(m, row, col)) {
                row += dr;
                col += dc;
            }
        }
        return new RangeObj(m, r.sheetName, row, col, row, col);
    }

    private RangeObj usedRange(SpreadsheetModel m, String sheet) {
        int r0 = Integer.MAX_VALUE;
        int c0 = Integer.MAX_VALUE;
        int r1 = -1;
        int c1 = -1;
        for (int[] p : m.storedCells()) {
            if (m.getInput(p[0], p[1]).isEmpty()) continue;
            r0 = Math.min(r0, p[0]);
            c0 = Math.min(c0, p[1]);
            r1 = Math.max(r1, p[0]);
            c1 = Math.max(c1, p[1]);
        }
        if (r1 < 0) return new RangeObj(m, sheet, 0, 0, 0, 0);
        return new RangeObj(m, sheet, r0, c0, r1, c1);
    }

    private static String colLetters(int c) {
        return FormulaEngine.columnLabel(c);
    }

    private String address(RangeObj r) {
        String a = "$" + colLetters(r.c0) + "$" + (r.r0 + 1);
        if (r.single()) return a;
        return a + ":$" + colLetters(r.c1) + "$" + (r.r1 + 1);
    }

    private SheetObj sheetObj(String name) throws MacroException {
        if (workbook == null) {
            if (name.equalsIgnoreCase(activeName)) return new SheetObj(model, activeName);
            throw err("Subscript out of range: sheet '" + name + "' does not exist");
        }
        for (String n : workbook.sheetNames()) {
            if (n.equalsIgnoreCase(name)) {
                SpreadsheetModel m = workbook.sheet(n);
                if (m != null) return new SheetObj(m, n);
            }
        }
        throw err("Subscript out of range: sheet '" + name + "' does not exist");
    }

    private SheetObj active() {
        return new SheetObj(model, activeName);
    }

    // ------------------------------------------------------------------
    // Root names: Range, Cells, ActiveSheet, built-in functions ...
    // ------------------------------------------------------------------

    private Object root(String name, List<Object> a, boolean parens, Frame f) throws MacroException {
        SheetObj sh = active();
        switch (name) {
            case "range": return parseRange(a, sh.m, sh.name);
            case "cells": return cellsOf(null, sh.m, sh.name, a);
            case "rows": case "columns": {
                RangeObj all = checked(new RangeObj(sh.m, sh.name, 0, 0, sh.m.getRowCount() - 1, sh.m.getColumnCount() - 1));
                DimObj d = new DimObj(all, name.equals("rows"));
                return a.isEmpty() ? d : dimItem(d, a.get(0));
            }
            case "activecell": {
                int[] s = workbook == null ? null : workbook.selection();
                return s == null ? checked(new RangeObj(sh.m, sh.name, 0, 0, 0, 0)) : checked(new RangeObj(sh.m, sh.name, s[0], s[1], s[0], s[1]));
            }
            case "selection": {
                int[] s = workbook == null ? null : workbook.selection();
                return s == null ? checked(new RangeObj(sh.m, sh.name, 0, 0, 0, 0)) : checked(new RangeObj(sh.m, sh.name, s[0], s[1], s[2], s[3]));
            }
            case "activesheet": return sh;
            case "thisworkbook": case "activeworkbook": return new AppObj();
            case "application": return new AppObj();
            case "worksheets": case "sheets": {
                if (a.isEmpty()) return new AppObj();
                Object k = a.get(0);
                if (k instanceof Double) {
                    int i = (int) Math.round((Double) k) - 1;
                    List<String> names = workbook == null ? java.util.Collections.singletonList(activeName) : workbook.sheetNames();
                    if (i < 0 || i >= names.size()) throw err("Subscript out of range");
                    return sheetObj(names.get(i));
                }
                return sheetObj(str(k));
            }
            case "usedrange": return usedRange(sh.m, sh.name);
            case "err": return errObj;
            case "msgbox": {
                if (a.isEmpty()) throw err("MsgBox needs a message");
                showMsg(str(a.get(0)));
                int btn = a.size() > 1 ? (int) toNum(a.get(1)) & 15 : 0;
                return btn == 4 || btn == 3 ? 6.0 : 1.0;
            }
            case "inputbox": {
                if (a.isEmpty()) throw err("InputBox needs a prompt");
                String r = inputBox == null ? null : inputBox.apply(str(a.get(0)));
                return r == null ? "" : r;
            }
            case "evaluate": {
                if (a.size() != 1) throw err("Evaluate needs one argument");
                String expr = str(a.get(0)).trim();
                if (!expr.startsWith("=")) expr = "=" + expr;
                SpreadsheetModel tmp = sh.m;
                String formula = expr;
                // evaluate in a scratch copy so no cell changes: use the model's engine through a hidden cell
                return evaluateFormula(tmp, formula);
            }
            default:
        }
        Double constant = CONSTANTS.get(name);
        if (constant != null) return constant;
        switch (name) {
            case "vbcrlf": return "\r\n";
            case "vblf": case "vbnewline": return "\n";
            case "vbcr": return "\r";
            case "vbtab": return "\t";
            case "vbnullstring": return "";
            case "pi": return Math.PI;
            case "rnd": return random.nextDouble();
            case "now": return LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
            case "date": return LocalDate.now().toString();
            case "time": return LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
            case "timer": return (double) (System.currentTimeMillis() % 86_400_000L) / 1000.0;
            default:
        }
        return builtin(name, a, f);
    }

    private Object evaluateFormula(SpreadsheetModel m, String formula) throws MacroException {
        // Use a far-away scratch cell in a throw-away way: set, read, restore.
        int r = m.getRowCount() - 1;
        int c = m.getColumnCount() - 1;
        String old = m.getInput(r, c);
        try {
            m.setInput(r, c, formula);
            Object v = m.getComputedValue(r, c);
            return v == null ? "" : v;
        } finally {
            m.setInput(r, c, old);
        }
    }

    private Object dimItem(DimObj d, Object idx) throws MacroException {
        int i = (int) Math.round(toNum(idx)) - 1;
        RangeObj r = d.r;
        if (d.rows) return checked(new RangeObj(r.m, r.sheetName, r.r0 + i, r.c0, r.r0 + i, r.c1));
        return checked(new RangeObj(r.m, r.sheetName, r.r0, r.c0 + i, r.r1, r.c0 + i));
    }

    // ------------------------------------------------------------------
    // Object members
    // ------------------------------------------------------------------

    private Object member(Object obj, String name, List<Object> a, boolean set, Object value) throws MacroException {
        if (obj instanceof RangeObj) return rangeMember((RangeObj) obj, name, a, set, value);
        if (obj instanceof FontObj) return fontMember((FontObj) obj, name, set, value);
        if (obj instanceof InteriorObj) {
            if (!name.equals("color") && !name.equals("colorindex")) throw err("Unsupported property: Interior." + name);
            RangeObj r = ((InteriorObj) obj).r;
            if (set) {
                Color c = fromVba(toNum(value));
                styleEach(r, st -> st.background = c);
                return null;
            }
            CellStyle st = r.m.getStyle(r.r0, r.c0);
            return toVba(st == null ? null : st.background);
        }
        if (obj instanceof DimObj) {
            DimObj d = (DimObj) obj;
            switch (name) {
                case "count": return (double) (d.rows ? d.r.r1 - d.r.r0 + 1 : d.r.c1 - d.r.c0 + 1);
                case "item": return dimItem(d, a.get(0));
                case "entirerow": case "entirecolumn": return d.r;
                default: throw err("Unsupported property: " + name);
            }
        }
        if (obj instanceof SheetObj) return sheetMember((SheetObj) obj, name, a, set, value);
        if (obj instanceof AppObj) return appMember(name, a, set, value);
        if (obj instanceof WfObj) return worksheetFunction(name, a);
        if (obj instanceof ErrObj) {
            switch (name) {
                case "number": return (double) errNumber;
                case "description": return errDesc;
                case "clear":
                    errNumber = 0;
                    errDesc = "";
                    return null;
                case "raise":
                    throw err(a.size() > 2 ? str(a.get(2)) : "Run-time error " + (a.isEmpty() ? "" : str(a.get(0))));
                default: throw err("Unsupported property: Err." + name);
            }
        }
        if (obj instanceof VArray) {
            if (name.equals("count") || name.equals("length")) return (double) ((VArray) obj).data.length;
            throw err("Unsupported array member: " + name);
        }
        if (obj instanceof String) {
            if (name.equals("length")) return (double) ((String) obj).length();
        }
        throw err("Object doesn't support '" + name + "'");
    }

    private Object fontMember(FontObj fo, String name, boolean set, Object value) throws MacroException {
        RangeObj r = fo.r;
        if (set) {
            switch (name) {
                case "bold": {
                    boolean b = toBool(value);
                    styleEach(r, st -> st.bold = b);
                    break;
                }
                case "italic": {
                    boolean b = toBool(value);
                    styleEach(r, st -> st.italic = b);
                    break;
                }
                case "size": {
                    int s = (int) toNum(value);
                    styleEach(r, st -> st.fontSize = s);
                    break;
                }
                case "color": {
                    Color c = fromVba(toNum(value));
                    styleEach(r, st -> st.foreground = c);
                    break;
                }
                case "name": {
                    String n = str(value);
                    styleEach(r, st -> st.fontFamily = n.isEmpty() ? null : n);
                    break;
                }
                default: throw err("Unsupported property: Font." + name);
            }
            return null;
        }
        CellStyle st = r.m.getStyle(r.r0, r.c0);
        switch (name) {
            case "bold": return st != null && st.bold;
            case "italic": return st != null && st.italic;
            case "size": return (double) (st == null || st.fontSize <= 0 ? 11 : st.fontSize);
            case "color": return toVba(st == null || st.foreground == null ? Color.BLACK : st.foreground) == 16777215 ? 0.0 : toVba(st.foreground);
            case "name": return st == null || st.fontFamily == null ? "Calibri" : st.fontFamily;
            default: throw err("Unsupported property: Font." + name);
        }
    }

    private Object rangeMember(RangeObj r, String name, List<Object> a, boolean set, Object value) throws MacroException {
        switch (name) {
            case "value": case "value2": case "formula": case "formular1c1": {
                if (set) {
                    writeRange(r, value, name.equals("formula"));
                    return null;
                }
                if (name.equals("formula") && r.single()) return r.m.getInput(r.r0, r.c0);
                return rangeValue(r);
            }
            case "text": return String.valueOf(r.m.getValueAt(r.r0, r.c0));
            case "address": return address(r);
            case "row": return (double) (r.r0 + 1);
            case "column": return (double) (r.c0 + 1);
            case "count": return (double) r.count();
            case "rows": return new DimObj(r, true);
            case "columns": return new DimObj(r, false);
            case "cells": return cellsOf(r, r.m, r.sheetName, a);
            case "item": return cellsOf(r, r.m, r.sheetName, a);
            case "range": return parseRangeIn(r, a);
            case "offset": {
                int dr = a.size() > 0 && a.get(0) != null ? (int) Math.round(toNum(a.get(0))) : 0;
                int dc = a.size() > 1 && a.get(1) != null ? (int) Math.round(toNum(a.get(1))) : 0;
                return checked(new RangeObj(r.m, r.sheetName, r.r0 + dr, r.c0 + dc, r.r1 + dr, r.c1 + dc));
            }
            case "resize": {
                int nr = a.size() > 0 && a.get(0) != null ? (int) Math.round(toNum(a.get(0))) : r.r1 - r.r0 + 1;
                int nc = a.size() > 1 && a.get(1) != null ? (int) Math.round(toNum(a.get(1))) : r.c1 - r.c0 + 1;
                if (nr < 1 || nc < 1) throw err("Invalid procedure call: Resize needs positive sizes");
                return checked(new RangeObj(r.m, r.sheetName, r.r0, r.c0, r.r0 + nr - 1, r.c0 + nc - 1));
            }
            case "end": return endOf(r, a.isEmpty() ? 0 : (int) Math.round(toNum(a.get(0))));
            case "entirerow": return checked(new RangeObj(r.m, r.sheetName, r.r0, 0, r.r1, r.m.getColumnCount() - 1));
            case "entirecolumn": return checked(new RangeObj(r.m, r.sheetName, 0, r.c0, r.m.getRowCount() - 1, r.c1));
            case "font": return new FontObj(r);
            case "interior": return new InteriorObj(r);
            case "parent": return new SheetObj(r.m, r.sheetName);
            case "horizontalalignment": {
                if (set) {
                    int code = (int) toNum(value);
                    int align = code == -4108 ? SwingConstants.CENTER : code == -4152 ? SwingConstants.RIGHT
                            : code == -4131 ? SwingConstants.LEFT : CellStyle.ALIGN_AUTO;
                    styleEach(r, st -> st.alignment = align);
                    return null;
                }
                CellStyle st = r.m.getStyle(r.r0, r.c0);
                int al = st == null ? CellStyle.ALIGN_AUTO : st.alignment;
                return al == SwingConstants.CENTER ? -4108.0 : al == SwingConstants.RIGHT ? -4152.0 : al == SwingConstants.LEFT ? -4131.0 : 1.0;
            }
            case "numberformat": {
                if (!set) return "General";
                applyNumberFormat(r, str(value));
                return null;
            }
            case "columnwidth": {
                if (set) {
                    int px = (int) Math.max(10, Math.min(1000, toNum(value) * 7));
                    for (int c = r.c0; c <= r.c1; c++) r.m.setColumnWidth(c, px);
                    return null;
                }
                return (double) r.m.getColumnWidth(r.c0) / 7.0;
            }
            case "clearcontents": case "clear": {
                for (int i = r.r0; i <= r.r1; i++) for (int j = r.c0; j <= r.c1; j++) r.m.clearContents(i, j);
                if (name.equals("clear")) styleEach(r, CellStyle::reset);
                return null;
            }
            case "select": case "activate": case "autofit": case "wraptext": case "verticalalignment": case "locked": case "hidden":
            case "rowheight": case "formulahidden": case "numberformatlocal":
                return null;
            default:
                throw err("Unsupported property or method: " + name);
        }
    }

    private RangeObj parseRangeIn(RangeObj base, List<Object> a) throws MacroException {
        RangeObj rel = parseRange(a, base.m, base.sheetName);
        return checked(new RangeObj(base.m, base.sheetName, base.r0 + rel.r0, base.c0 + rel.c0, base.r0 + rel.r1, base.c0 + rel.c1));
    }

    private void applyNumberFormat(RangeObj r, String fmt) {
        String f = fmt.trim();
        int kind = CellFormats.GENERAL;
        int dec = -1;
        String low = f.toLowerCase(Locale.ROOT);
        if (low.equals("general") || low.isEmpty()) {
            kind = CellFormats.GENERAL;
        } else {
            int dot = f.indexOf('.');
            if (dot >= 0) {
                int n = 0;
                for (int i = dot + 1; i < f.length() && (f.charAt(i) == '0' || f.charAt(i) == '#'); i++) n++;
                dec = n;
            } else {
                dec = 0;
            }
            if (f.contains("%")) kind = CellFormats.PERCENT;
            else if (f.contains("$") || f.contains("€") || f.contains("£") || low.contains("currency")) kind = CellFormats.CURRENCY;
            else if (low.contains("e+") || low.contains("e-")) kind = CellFormats.SCIENTIFIC;
            else if (f.contains(",")) kind = CellFormats.COMMA;
            else kind = CellFormats.NUMBER;
        }
        final int k = kind;
        final int d = dec;
        styleEach(r, st -> {
            st.numFmt = k;
            st.decimals = k == CellFormats.GENERAL ? -1 : d;
        });
    }

    private Object sheetMember(SheetObj s, String name, List<Object> a, boolean set, Object value) throws MacroException {
        switch (name) {
            case "name": return s.name;
            case "range": return parseRange(a, s.m, s.name);
            case "cells": return cellsOf(null, s.m, s.name, a);
            case "usedrange": return usedRange(s.m, s.name);
            case "rows": return new DimObj(checked(new RangeObj(s.m, s.name, 0, 0, s.m.getRowCount() - 1, s.m.getColumnCount() - 1)), true);
            case "columns": return new DimObj(checked(new RangeObj(s.m, s.name, 0, 0, s.m.getRowCount() - 1, s.m.getColumnCount() - 1)), false);
            case "activate": case "select": return null;
            case "visible": case "protect": case "unprotect": return null;
            default: throw err("Unsupported worksheet member: " + name);
        }
    }

    private Object appMember(String name, List<Object> a, boolean set, Object value) throws MacroException {
        switch (name) {
            case "screenupdating": case "displayalerts": case "calculation": case "enableevents": case "statusbar": case "cutcopymode":
                return set ? null : (Object) Boolean.TRUE;
            case "worksheetfunction": return new WfObj();
            case "activesheet": return active();
            case "worksheets": case "sheets": {
                if (a.isEmpty()) return new AppObj();
                Object k = a.get(0);
                if (k instanceof Double) {
                    List<String> names = workbook == null ? java.util.Collections.singletonList(activeName) : workbook.sheetNames();
                    int i = (int) Math.round((Double) k) - 1;
                    if (i < 0 || i >= names.size()) throw err("Subscript out of range");
                    return sheetObj(names.get(i));
                }
                return sheetObj(str(k));
            }
            case "count": {
                return (double) (workbook == null ? 1 : workbook.sheetNames().size());
            }
            case "name": return "Spreadsheet Pro";
            case "activecell": return root("activecell", a, false, null);
            case "selection": return root("selection", a, false, null);
            case "range": return root("range", a, true, null);
            case "cells": return root("cells", a, true, null);
            case "wait": case "run": case "calculate": case "quit": return null;
            default:
                // Application.Sum(...), Application.Max(...) etc.
                return worksheetFunction(name, a);
        }
    }

    private List<Double> numbersOf(List<Object> a) throws MacroException {
        List<Double> out = new ArrayList<>();
        for (Object o : a) collect(o, out);
        return out;
    }

    private void collect(Object o, List<Double> out) throws MacroException {
        if (o instanceof RangeObj) {
            RangeObj r = (RangeObj) o;
            for (int i = r.r0; i <= r.r1; i++) {
                for (int j = r.c0; j <= r.c1; j++) {
                    Object v = r.m.getComputedValue(i, j);
                    if (v instanceof Double) out.add((Double) v);
                }
            }
        } else if (o instanceof VArray) {
            for (Object x : ((VArray) o).data) if (x instanceof Double) out.add((Double) x);
        } else if (o instanceof Double) {
            out.add((Double) o);
        }
    }

    private Object worksheetFunction(String name, List<Object> a) throws MacroException {
        switch (name) {
            case "sum": {
                double s = 0;
                for (double d : numbersOf(a)) s += d;
                return s;
            }
            case "average": {
                List<Double> n = numbersOf(a);
                if (n.isEmpty()) throw err("Division by zero (no numbers to average)");
                double s = 0;
                for (double d : n) s += d;
                return s / n.size();
            }
            case "count": return (double) numbersOf(a).size();
            case "counta": {
                double n = 0;
                for (Object o : a) {
                    if (o instanceof RangeObj) {
                        RangeObj r = (RangeObj) o;
                        for (int i = r.r0; i <= r.r1; i++) for (int j = r.c0; j <= r.c1; j++) if (!r.m.getInput(i, j).isEmpty()) n++;
                    } else if (o != null && !str(o).isEmpty()) n++;
                }
                return n;
            }
            case "max": {
                List<Double> n = numbersOf(a);
                double m = n.isEmpty() ? 0 : n.get(0);
                for (double d : n) m = Math.max(m, d);
                return m;
            }
            case "min": {
                List<Double> n = numbersOf(a);
                double m = n.isEmpty() ? 0 : n.get(0);
                for (double d : n) m = Math.min(m, d);
                return m;
            }
            case "median": {
                List<Double> n = numbersOf(a);
                if (n.isEmpty()) throw err("No numbers for Median");
                java.util.Collections.sort(n);
                int k = n.size();
                return k % 2 == 1 ? n.get(k / 2) : (n.get(k / 2 - 1) + n.get(k / 2)) / 2;
            }
            case "product": {
                double p = 1;
                for (double d : numbersOf(a)) p *= d;
                return p;
            }
            case "round": return builtin("round", a, null);
            case "countif": {
                if (a.size() != 2 || !(a.get(0) instanceof RangeObj)) throw err("CountIf needs a range and a criterion");
                RangeObj r = (RangeObj) a.get(0);
                String crit = str(a.get(1));
                double n = 0;
                for (int i = r.r0; i <= r.r1; i++) {
                    for (int j = r.c0; j <= r.c1; j++) {
                        if (criterionMatches(r.m.getComputedValue(i, j), crit)) n++;
                    }
                }
                return n;
            }
            case "sumif": {
                if (a.size() < 2 || !(a.get(0) instanceof RangeObj)) throw err("SumIf needs a range and a criterion");
                RangeObj r = (RangeObj) a.get(0);
                RangeObj sum = a.size() > 2 && a.get(2) instanceof RangeObj ? (RangeObj) a.get(2) : r;
                String crit = str(a.get(1));
                double s = 0;
                for (int i = r.r0; i <= r.r1; i++) {
                    for (int j = r.c0; j <= r.c1; j++) {
                        if (criterionMatches(r.m.getComputedValue(i, j), crit)) {
                            Object v = sum.m.getComputedValue(sum.r0 + (i - r.r0), sum.c0 + (j - r.c0));
                            if (v instanceof Double) s += (Double) v;
                        }
                    }
                }
                return s;
            }
            default: throw err("Unsupported function: WorksheetFunction." + name);
        }
    }

    private boolean criterionMatches(Object cell, String crit) throws MacroException {
        String op = "=";
        String val = crit;
        for (String o : new String[]{"<>", "<=", ">=", "<", ">", "="}) {
            if (crit.startsWith(o)) {
                op = o;
                val = crit.substring(o.length());
                break;
            }
        }
        Object c = cell == null ? "" : cell;
        if (isNumeric(val) && !val.trim().isEmpty() && c instanceof Double) return compare(c, Double.parseDouble(val.trim()), op);
        if (op.equals("=") && (val.contains("*") || val.contains("?"))) return like(str(c), val);
        return compare(str(c), val, op);
    }

    // ------------------------------------------------------------------
    // Built-in functions
    // ------------------------------------------------------------------

    private void need(List<Object> a, int min, int max, String name) throws MacroException {
        if (a.size() < min || a.size() > max) throw err("Wrong number of arguments for " + name);
    }

    private Object builtin(String name, List<Object> a, Frame f) throws MacroException {
        switch (name) {
            case "int": need(a, 1, 1, "Int"); return Math.floor(toNum(a.get(0)));
            case "fix": { need(a, 1, 1, "Fix"); double x = toNum(a.get(0)); return x < 0 ? Math.ceil(x) : Math.floor(x); }
            case "abs": need(a, 1, 1, "Abs"); return Math.abs(toNum(a.get(0)));
            case "sgn": need(a, 1, 1, "Sgn"); return Math.signum(toNum(a.get(0)));
            case "sqr": case "sqrt": {
                need(a, 1, 1, "Sqr");
                double x = toNum(a.get(0));
                if (x < 0) throw err("Invalid procedure call: Sqr of a negative number");
                return Math.sqrt(x);
            }
            case "exp": need(a, 1, 1, "Exp"); return Math.exp(toNum(a.get(0)));
            case "log": {
                need(a, 1, 1, "Log");
                double x = toNum(a.get(0));
                if (x <= 0) throw err("Invalid procedure call: Log of zero or a negative number");
                return Math.log(x);
            }
            case "sin": need(a, 1, 1, "Sin"); return Math.sin(toNum(a.get(0)));
            case "cos": need(a, 1, 1, "Cos"); return Math.cos(toNum(a.get(0)));
            case "tan": need(a, 1, 1, "Tan"); return Math.tan(toNum(a.get(0)));
            case "atn": need(a, 1, 1, "Atn"); return Math.atan(toNum(a.get(0)));
            case "rnd": return random.nextDouble();
            case "round": {
                need(a, 1, 2, "Round");
                int digits = a.size() == 2 ? (int) toNum(a.get(1)) : 0;
                double x = toNum(a.get(0));
                if (Double.isNaN(x) || Double.isInfinite(x)) return x;
                return new BigDecimal(x).setScale(digits, RoundingMode.HALF_EVEN).doubleValue();
            }
            case "len": need(a, 1, 1, "Len"); return (double) str(a.get(0)).length();
            case "ucase": need(a, 1, 1, "UCase"); return str(a.get(0)).toUpperCase(Locale.ROOT);
            case "lcase": need(a, 1, 1, "LCase"); return str(a.get(0)).toLowerCase(Locale.ROOT);
            case "trim": need(a, 1, 1, "Trim"); return str(a.get(0)).replaceAll("^ +| +$", "");
            case "ltrim": need(a, 1, 1, "LTrim"); return str(a.get(0)).replaceAll("^\\s+", "");
            case "rtrim": need(a, 1, 1, "RTrim"); return str(a.get(0)).replaceAll("\\s+$", "");
            case "left": {
                need(a, 2, 2, "Left");
                String s = str(a.get(0));
                int n = (int) toNum(a.get(1));
                if (n < 0) throw err("Invalid procedure call: Left with a negative length");
                return s.substring(0, Math.min(n, s.length()));
            }
            case "right": {
                need(a, 2, 2, "Right");
                String s = str(a.get(0));
                int n = (int) toNum(a.get(1));
                if (n < 0) throw err("Invalid procedure call: Right with a negative length");
                return s.substring(Math.max(0, s.length() - n));
            }
            case "mid": {
                need(a, 2, 3, "Mid");
                String s = str(a.get(0));
                int st = (int) toNum(a.get(1)) - 1;
                if (st < 0) throw err("Invalid procedure call: Mid start must be 1 or more");
                if (st >= s.length()) return "";
                int end = a.size() == 3 ? st + Math.max(0, (int) toNum(a.get(2))) : s.length();
                return s.substring(st, Math.min(s.length(), end));
            }
            case "instr": {
                need(a, 2, 4, "InStr");
                int start = 1;
                int k = 0;
                if (a.get(0) instanceof Double && a.size() >= 3) {
                    start = (int) toNum(a.get(0));
                    k = 1;
                }
                String hay = str(a.get(k));
                String needle = str(a.get(k + 1));
                boolean text = a.size() > k + 2 && toNum(a.get(k + 2)) == 1;
                if (text) {
                    hay = hay.toLowerCase(Locale.ROOT);
                    needle = needle.toLowerCase(Locale.ROOT);
                }
                return (double) (hay.indexOf(needle, Math.max(0, start - 1)) + 1);
            }
            case "instrrev": {
                need(a, 2, 3, "InStrRev");
                return (double) (str(a.get(0)).lastIndexOf(str(a.get(1))) + 1);
            }
            case "replace": {
                need(a, 3, 6, "Replace");
                String s = str(a.get(0));
                String find = str(a.get(1));
                if (find.isEmpty()) return s;
                boolean text = a.size() > 5 && toNum(a.get(5)) == 1;
                if (text) return Pattern.compile(Pattern.quote(find), CI).matcher(s).replaceAll(Matcher.quoteReplacement(str(a.get(2))));
                return s.replace(find, str(a.get(2)));
            }
            case "space": need(a, 1, 1, "Space"); return " ".repeat((int) Math.max(0, Math.min(32767, toNum(a.get(0)))));
            case "string": {
                need(a, 2, 2, "String");
                String ch = a.get(1) instanceof Double ? String.valueOf((char) (int) toNum(a.get(1))) : str(a.get(1));
                return ch.isEmpty() ? "" : ch.substring(0, 1).repeat((int) Math.max(0, Math.min(32767, toNum(a.get(0)))));
            }
            case "chr": need(a, 1, 1, "Chr"); return String.valueOf((char) (int) toNum(a.get(0)));
            case "asc": { need(a, 1, 1, "Asc"); String s = str(a.get(0)); if (s.isEmpty()) throw err("Invalid procedure call: Asc of an empty string"); return (double) s.charAt(0); }
            case "strreverse": need(a, 1, 1, "StrReverse"); return new StringBuilder(str(a.get(0))).reverse().toString();
            case "strcomp": {
                need(a, 2, 3, "StrComp");
                boolean text = a.size() == 3 && toNum(a.get(2)) == 1;
                int c = text ? str(a.get(0)).compareToIgnoreCase(str(a.get(1))) : str(a.get(0)).compareTo(str(a.get(1)));
                return (double) Integer.signum(c);
            }
            case "str": need(a, 1, 1, "Str"); return (toNum(a.get(0)) >= 0 ? " " : "") + str(a.get(0));
            case "cstr": need(a, 1, 1, "CStr"); return str(a.get(0));
            case "val": {
                need(a, 1, 1, "Val");
                Matcher m = LEADING_NUMBER.matcher(str(a.get(0)));
                return m.find() ? Double.parseDouble(m.group().trim()) : 0.0;
            }
            case "cint": case "clng": need(a, 1, 1, name); return (double) Math.round(toNum(a.get(0)));
            case "cdbl": case "csng": case "ccur": case "cdate": need(a, 1, 1, name); return toNum(a.get(0));
            case "cbool": need(a, 1, 1, "CBool"); return toBool(a.get(0));
            case "isnumeric": {
                need(a, 1, 1, "IsNumeric");
                Object v = a.get(0);
                if (v instanceof RangeObj) v = rangeValue((RangeObj) v);
                return v instanceof Double || (v instanceof String && !((String) v).trim().isEmpty() && isNumeric((String) v));
            }
            case "isempty": {
                need(a, 1, 1, "IsEmpty");
                Object v = a.get(0);
                if (v instanceof RangeObj) v = rangeValue((RangeObj) v);
                return v == null || (v instanceof String && ((String) v).isEmpty());
            }
            case "isnull": need(a, 1, 1, "IsNull"); return false;
            case "isarray": need(a, 1, 1, "IsArray"); return a.get(0) instanceof VArray;
            case "isdate": need(a, 1, 1, "IsDate"); return a.get(0) instanceof String && ((String) a.get(0)).matches("\\d{4}-\\d{2}-\\d{2}.*");
            case "iserror": need(a, 1, 1, "IsError"); return false;
            case "isobject": need(a, 1, 1, "IsObject"); return a.get(0) instanceof RangeObj || a.get(0) instanceof SheetObj;
            case "typename": {
                need(a, 1, 1, "TypeName");
                Object v = a.get(0);
                return v == null ? "Empty" : v instanceof Double ? "Double" : v instanceof String ? "String" : v instanceof Boolean ? "Boolean"
                        : v instanceof RangeObj ? "Range" : v instanceof VArray ? "Variant()" : "Object";
            }
            case "year": return dateField(a, 0);
            case "month": return dateField(a, 1);
            case "day": return dateField(a, 2);
            case "weekday": {
                need(a, 1, 2, "Weekday");
                LocalDate d = parseDate(str(a.get(0)));
                return (double) (d.getDayOfWeek().getValue() % 7 + 1);
            }
            case "dateadd": {
                need(a, 3, 3, "DateAdd");
                LocalDate d = parseDate(str(a.get(2)));
                int n = (int) toNum(a.get(1));
                switch (str(a.get(0)).toLowerCase(Locale.ROOT)) {
                    case "d": return d.plusDays(n).toString();
                    case "m": return d.plusMonths(n).toString();
                    case "yyyy": return d.plusYears(n).toString();
                    case "ww": return d.plusWeeks(n).toString();
                    default: throw err("Unsupported DateAdd interval");
                }
            }
            case "datediff": {
                need(a, 3, 3, "DateDiff");
                LocalDate d1 = parseDate(str(a.get(1)));
                LocalDate d2 = parseDate(str(a.get(2)));
                switch (str(a.get(0)).toLowerCase(Locale.ROOT)) {
                    case "d": return (double) java.time.temporal.ChronoUnit.DAYS.between(d1, d2);
                    case "m": return (double) java.time.temporal.ChronoUnit.MONTHS.between(d1, d2);
                    case "yyyy": return (double) java.time.temporal.ChronoUnit.YEARS.between(d1, d2);
                    case "ww": return (double) java.time.temporal.ChronoUnit.WEEKS.between(d1, d2);
                    default: throw err("Unsupported DateDiff interval");
                }
            }
            case "dateserial": {
                need(a, 3, 3, "DateSerial");
                return LocalDate.of((int) toNum(a.get(0)), 1, 1).plusMonths((long) toNum(a.get(1)) - 1).plusDays((long) toNum(a.get(2)) - 1).toString();
            }
            case "rgb": {
                need(a, 3, 3, "RGB");
                int r = (int) toNum(a.get(0)) & 255;
                int g = (int) toNum(a.get(1)) & 255;
                int b = (int) toNum(a.get(2)) & 255;
                return (double) (r | (g << 8) | (b << 16));
            }
            case "format": return format(a);
            case "formatnumber": case "formatcurrency": case "formatpercent": {
                need(a, 1, 2, name);
                int dec = a.size() == 2 ? (int) toNum(a.get(1)) : 2;
                double x = toNum(a.get(0));
                String p = dec > 0 ? "#,##0." + "0".repeat(dec) : "#,##0";
                if (name.equals("formatpercent")) return new java.text.DecimalFormat(p + "%").format(x);
                return (name.equals("formatcurrency") ? "$" : "") + new java.text.DecimalFormat(p).format(x);
            }
            case "split": {
                need(a, 1, 3, "Split");
                String s = str(a.get(0));
                String d = a.size() > 1 ? str(a.get(1)) : " ";
                String[] parts = d.isEmpty() ? new String[]{s} : s.split(Pattern.quote(d), -1);
                VArray arr = new VArray(new int[]{0}, new int[]{parts.length - 1});
                System.arraycopy(parts, 0, arr.data, 0, parts.length);
                return arr;
            }
            case "join": {
                need(a, 1, 2, "Join");
                if (!(a.get(0) instanceof VArray)) throw err("Join needs an array");
                String d = a.size() > 1 ? str(a.get(1)) : " ";
                StringBuilder sb = new StringBuilder();
                Object[] data = ((VArray) a.get(0)).data;
                for (int i = 0; i < data.length; i++) {
                    if (i > 0) sb.append(d);
                    sb.append(str(data[i]));
                }
                return sb.toString();
            }
            case "array": {
                VArray arr = new VArray(new int[]{0}, new int[]{a.size() - 1});
                for (int i = 0; i < a.size(); i++) arr.data[i] = a.get(i);
                return arr;
            }
            case "ubound": case "lbound": {
                need(a, 1, 2, name);
                if (!(a.get(0) instanceof VArray)) throw err("Type mismatch: UBound/LBound needs an array");
                VArray v = (VArray) a.get(0);
                int dim = a.size() == 2 ? (int) toNum(a.get(1)) - 1 : 0;
                if (dim < 0 || dim >= v.lo.length) throw err("Subscript out of range");
                return (double) (name.equals("ubound") ? v.hi[dim] : v.lo[dim]);
            }
            case "max": case "min": {
                if (a.isEmpty()) throw err("Wrong number of arguments for " + name);
                return worksheetFunction(name, a);
            }
            default:
                throw err("Variable or procedure not defined: " + name);
        }
    }

    private LocalDate parseDate(String s) throws MacroException {
        try {
            return LocalDate.parse(s.trim().substring(0, Math.min(10, s.trim().length())));
        } catch (RuntimeException e) {
            throw err("Type mismatch: '" + s + "' is not a date (use yyyy-mm-dd)");
        }
    }

    private Object dateField(List<Object> a, int which) throws MacroException {
        need(a, 1, 1, "Year/Month/Day");
        LocalDate d = parseDate(str(a.get(0)));
        return (double) (which == 0 ? d.getYear() : which == 1 ? d.getMonthValue() : d.getDayOfMonth());
    }

    private Object format(List<Object> a) throws MacroException {
        need(a, 1, 2, "Format");
        Object v = a.get(0);
        if (a.size() == 1) return str(v);
        String fmt = str(a.get(1));
        String low = fmt.toLowerCase(Locale.ROOT);
        try {
            if (low.contains("yy") || low.contains("dd") || (low.contains("mm") && low.contains("-")) || low.contains("hh")) {
                String s = str(v);
                LocalDateTime t = s.length() > 10 ? LocalDateTime.parse(s.replace(' ', 'T')) : LocalDate.parse(s).atStartOfDay();
                boolean ampm = low.contains("am/pm");
                String j = fmt.replaceAll("(?i)(?<=[:h])mm", "\u0001").replace("mmmm", "MMMM").replace("mmm", "MMM").replace("mm", "MM")
                        .replace("\u0001", "mm").replace("nn", "mm").replaceAll("(?i)am/pm", "a");
                if (!ampm) j = j.replace("hh", "HH");
                return t.format(DateTimeFormatter.ofPattern(j, Locale.ENGLISH));
            }
            double x = toNum(v);
            switch (low) {
                case "percent": return new java.text.DecimalFormat("0.00%").format(x);
                case "currency": return "$" + new java.text.DecimalFormat("#,##0.00").format(x);
                case "fixed": return new java.text.DecimalFormat("0.00").format(x);
                case "standard": return new java.text.DecimalFormat("#,##0.00").format(x);
                case "scientific": return new java.text.DecimalFormat("0.00E00").format(x);
                case "yes/no": return x != 0 ? "Yes" : "No";
                case "true/false": return x != 0 ? "True" : "False";
                case "on/off": return x != 0 ? "On" : "Off";
                default:
            }
            return new java.text.DecimalFormat(fmt.replace("\"", "'")).format(x);
        } catch (RuntimeException e) {
            throw err("Invalid format string: " + fmt);
        }
    }
}
