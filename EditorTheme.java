import java.awt.Color;

/** Colour scheme of the macro code editor. */
public final class EditorTheme {
    public final String name;
    public final boolean dark;
    public final Color bg, fg, caret, selection, currentLine, gutterBg, gutterFg, gutterCurrent, guide, bracket, match, errorLine, border;
    /** Colours indexed by {@link VbaSyntax} token kind. */
    public final Color[] kind = new Color[VbaSyntax.KINDS];

    private EditorTheme(String name, boolean dark, int bg, int fg, int caret, int selection, int currentLine, int gutterBg,
                        int gutterFg, int gutterCurrent, int keyword, int string, int comment, int number, int builtin,
                        int type, int object, int constant, int member) {
        this.name = name;
        this.dark = dark;
        this.bg = new Color(bg);
        this.fg = new Color(fg);
        this.caret = new Color(caret);
        this.selection = new Color(selection);
        this.currentLine = new Color(currentLine);
        this.gutterBg = new Color(gutterBg);
        this.gutterFg = new Color(gutterFg);
        this.gutterCurrent = new Color(gutterCurrent);
        Color f = this.fg;
        this.guide = new Color(f.getRed(), f.getGreen(), f.getBlue(), 38);
        this.bracket = new Color(0xFFD54F);
        this.match = dark ? new Color(0x6B5B00) : new Color(0xFFF176);
        this.errorLine = dark ? new Color(0x5A1D1D) : new Color(0xFFE0E0);
        this.border = dark ? new Color(0x3C3C3C) : new Color(0xD0D0D0);
        kind[VbaSyntax.PLAIN] = this.fg;
        kind[VbaSyntax.KEYWORD] = new Color(keyword);
        kind[VbaSyntax.STRING] = new Color(string);
        kind[VbaSyntax.COMMENT] = new Color(comment);
        kind[VbaSyntax.NUMBER] = new Color(number);
        kind[VbaSyntax.BUILTIN] = new Color(builtin);
        kind[VbaSyntax.OBJECT] = new Color(object);
        kind[VbaSyntax.TYPE] = new Color(type);
        kind[VbaSyntax.CONSTANT] = new Color(constant);
        kind[VbaSyntax.MEMBER] = new Color(member);
        kind[VbaSyntax.OPERATOR] = this.fg;
    }

    public static final EditorTheme[] ALL = {
            new EditorTheme("Dark+", true, 0x1E1E1E, 0xD4D4D4, 0xAEAFAD, 0x264F78, 0x282828, 0x1E1E1E, 0x6E7681, 0xC6C6C6,
                    0x569CD6, 0xCE9178, 0x6A9955, 0xB5CEA8, 0xDCDCAA, 0x4EC9B0, 0x9CDCFE, 0x4FC1FF, 0x9CDCFE),
            new EditorTheme("Light+", false, 0xFFFFFF, 0x1F1F1F, 0x000000, 0xADD6FF, 0xF3F6FB, 0xFFFFFF, 0x9AA0A6, 0x333333,
                    0x0000FF, 0xA31515, 0x008000, 0x098658, 0x795E26, 0x267F99, 0x001080, 0x0070C1, 0x001080),
            new EditorTheme("Monokai", true, 0x272822, 0xF8F8F2, 0xF8F8F0, 0x49483E, 0x32332B, 0x272822, 0x75715E, 0xC2C2BA,
                    0xF92672, 0xE6DB74, 0x75715E, 0xAE81FF, 0xA6E22E, 0x66D9EF, 0xFD971F, 0xAE81FF, 0xF8F8F2),
            new EditorTheme("Dracula", true, 0x282A36, 0xF8F8F2, 0xF8F8F0, 0x44475A, 0x303341, 0x282A36, 0x6272A4, 0xF8F8F2,
                    0xFF79C6, 0xF1FA8C, 0x6272A4, 0xBD93F9, 0x50FA7B, 0x8BE9FD, 0xFFB86C, 0xBD93F9, 0xF8F8F2),
            new EditorTheme("Solarized Light", false, 0xFDF6E3, 0x586E75, 0x586E75, 0xD6CDB0, 0xF3ECD5, 0xFDF6E3, 0x93A1A1, 0x586E75,
                    0x859900, 0x2AA198, 0x93A1A1, 0xD33682, 0xB58900, 0xCB4B16, 0x268BD2, 0x6C71C4, 0x268BD2),
            new EditorTheme("Spreadsheet Green", false, 0xFAFDFB, 0x1B2B22, 0x107C41, 0xBFE3CD, 0xEAF5EE, 0xF1F8F4, 0x8AA596, 0x107C41,
                    0x107C41, 0xB4531B, 0x7A8F82, 0x1F6FB5, 0x7B3FA0, 0x0F7C8C, 0x2B4A6F, 0x1F6FB5, 0x2B4A6F),
    };

    public static EditorTheme byName(String n) {
        for (EditorTheme t : ALL) if (t.name.equals(n)) return t;
        return null;
    }

    /** The theme that matches the application's light/dark mode. */
    public static EditorTheme forApp() {
        return Theme.dark() ? ALL[0] : ALL[1];
    }

    @Override
    public String toString() {
        return name;
    }
}
