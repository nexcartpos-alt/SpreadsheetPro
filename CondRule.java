import java.awt.Color;

/** A conditional-formatting rule: cells in a range that match get a fill / text colour (immutable). */
public final class CondRule {

    public enum Type { GREATER, LESS, EQUAL, BETWEEN, CONTAINS, DUPLICATE, COLOR_SCALE, ICON_SET, DATA_BAR }

    public final Type type;
    public final int r1;
    public final int c1;
    public final int r2;
    public final int c2;
    public final String v1;
    public final String v2;
    public final Color bg;
    /** Text colour, or null to keep the cell's own. */
    public final Color fg;

    public CondRule(Type type, int r1, int c1, int r2, int c2, String v1, String v2, Color bg, Color fg) {
        this.type = type;
        this.r1 = Math.min(r1, r2);
        this.c1 = Math.min(c1, c2);
        this.r2 = Math.max(r1, r2);
        this.c2 = Math.max(c1, c2);
        this.v1 = v1 == null ? "" : v1;
        this.v2 = v2 == null ? "" : v2;
        this.bg = bg;
        this.fg = fg;
    }

    public boolean isVisual() {
        return type == Type.COLOR_SCALE || type == Type.ICON_SET || type == Type.DATA_BAR;
    }

    /** Colour scale: 2 or 3 colours as hex, low to high (kept in v1, comma separated). */
    public static CondRule colorScale(int r1, int c1, int r2, int c2, String... hex) {
        return new CondRule(Type.COLOR_SCALE, r1, c1, r2, c2, String.join(",", hex), "", null, null);
    }

    public static CondRule iconSet(int r1, int c1, int r2, int c2, String set) {
        return new CondRule(Type.ICON_SET, r1, c1, r2, c2, set, "", null, null);
    }

    public static CondRule dataBar(int r1, int c1, int r2, int c2, String hex) {
        return new CondRule(Type.DATA_BAR, r1, c1, r2, c2, hex, "", null, null);
    }

    /** Number of icons in a set id (0 for an unknown id). */
    public static int iconCount(String set) {
        switch (set) {
            case "ARROWS3": case "TRAFFIC3": case "SYMBOLS3": case "FLAGS3": return 3;
            case "RATING4": case "ARROWS4": return 4;
            case "RATING5": case "ARROWS5": return 5;
            default: return 0;
        }
    }

    /** Parses "RRGGBB" (or null on bad input). */
    public static Color hex(String h) {
        try {
            return new Color(Integer.parseInt(h.trim(), 16) & 0xFFFFFF);
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** What a visual rule paints on one cell. */
    public static final class Visual {
        public Color bg;
        public int icon = -1;     // index within the set, higher = better
        public String iconSet;
        public double bar = -1;   // 0..1 bar length, -1 = none
        public Color barColor;
    }

    public boolean covers(int r, int c) {
        return r >= r1 && r <= r2 && c >= c1 && c <= c2;
    }

    public CondRule withRange(int nr1, int nc1, int nr2, int nc2) {
        return new CondRule(type, nr1, nc1, nr2, nc2, v1, v2, bg, fg);
    }

    /** The colour presets offered in the dialog (name, fill, text). */
    public static final String[] PRESET_NAMES = {
            "Light red fill with dark red text", "Yellow fill with dark yellow text",
            "Green fill with dark green text", "Light blue fill", "Orange fill"};
    public static final Color[] PRESET_BG = {
            new Color(0xFFC7CE), new Color(0xFFEB9C), new Color(0xC6EFCE), new Color(0xBDD7EE), new Color(0xF8CBAD)};
    public static final Color[] PRESET_FG = {
            new Color(0x9C0006), new Color(0x9C5700), new Color(0x006100), null, null};
}
