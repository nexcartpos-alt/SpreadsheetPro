import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.GraphicsEnvironment;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Fonts for the code editor: two bundled monospaced families plus every monospaced font on the system. */
public final class CodeFonts {
    private CodeFonts() {
    }

    private static final String[] BUNDLED = {
            "/fonts/DejaVuSansMono.ttf", "/fonts/DejaVuSansMono-Bold.ttf",
            "/fonts/LiberationMono-Regular.ttf", "/fonts/LiberationMono-Bold.ttf"};
    private static boolean registered;
    private static List<String> mono;

    /** Registers the bundled fonts with Java (once). */
    public static synchronized void register() {
        if (registered) return;
        registered = true;
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        for (String res : BUNDLED) {
            try (InputStream in = CodeFonts.class.getResourceAsStream(res)) {
                if (in == null) continue;
                ge.registerFont(Font.createFont(Font.TRUETYPE_FONT, in));
            } catch (Exception | Error e) {
                // a missing font is not fatal: the editor falls back to the logical "Monospaced" font
            }
        }
    }

    /** Installed families whose letters all have the same width. */
    public static synchronized List<String> monospaced() {
        if (mono != null) return mono;
        register();
        List<String> out = new ArrayList<>();
        Set<String> seen = new HashSet<>();
        BufferedImage img = new BufferedImage(1, 1, BufferedImage.TYPE_INT_RGB);
        java.awt.Graphics2D g = img.createGraphics();
        try {
            for (String fam : GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames()) {
                if (!seen.add(fam)) continue;
                try {
                    Font f = new Font(fam, Font.PLAIN, 14);
                    if (f.getFamily().equals(Font.DIALOG) && !fam.equals(Font.DIALOG)) continue;
                    FontMetrics fm = g.getFontMetrics(f);
                    int w = fm.charWidth('i');
                    if (w > 0 && w == fm.charWidth('M') && w == fm.charWidth('W') && w == fm.charWidth('.')
                            && f.canDisplay('A') && !fam.startsWith("Lucida Sans Typewriter Oblique")) out.add(fam);
                } catch (RuntimeException e) {
                    // skip fonts that cannot be measured
                }
            }
        } finally {
            g.dispose();
        }
        if (!out.contains("Monospaced")) out.add("Monospaced");
        java.util.Collections.sort(out, String.CASE_INSENSITIVE_ORDER);
        mono = out;
        return out;
    }

    /** The default family: the best coding font that is available. */
    public static String defaultFamily() {
        List<String> have = monospaced();
        for (String pref : Arrays.asList("JetBrains Mono", "Cascadia Code", "Cascadia Mono", "Consolas", "Fira Code",
                "Source Code Pro", "Menlo", "DejaVu Sans Mono", "Liberation Mono")) {
            if (have.contains(pref)) return pref;
        }
        return "Monospaced";
    }
}
