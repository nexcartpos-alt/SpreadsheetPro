import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.Insets;
import java.awt.Point;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.Window;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseWheelEvent;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.DefaultListCellRenderer;
import javax.swing.DefaultListModel;
import javax.swing.JComponent;
import javax.swing.JList;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.JWindow;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.event.CaretEvent;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.UndoableEditEvent;
import javax.swing.text.AbstractDocument;
import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultHighlighter;
import javax.swing.text.Element;
import javax.swing.text.Highlighter;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.Style;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyleContext;
import javax.swing.text.StyledDocument;
import javax.swing.undo.CompoundEdit;
import javax.swing.undo.UndoManager;
import javax.swing.undo.UndoableEdit;

/**
 * Code editor for macros: syntax colours, line-number gutter, auto-indent and auto-closing blocks,
 * completion popup, bracket matching, indent guides, find highlights, error squiggle, themes and fonts.
 */
public class CodeEditor extends JTextPane {

    public static final String UNIT = "    ";

    private EditorTheme theme = EditorTheme.forApp();
    private String family = "Monospaced";
    private int fontSize = 14;
    private final SimpleAttributeSet[] attrs = new SimpleAttributeSet[VbaSyntax.KINDS];
    private final Gutter gutter = new Gutter();
    private final UndoManager undo = new UndoManager();
    private CompoundEdit current;
    private long lastEditTime;
    private boolean explicitGroup;
    private boolean boundaryNext;
    private boolean highlighting;
    private boolean applying;
    private int errorLine = -1;
    private Object bracketA;
    private Object bracketB;
    private final List<Object> matchTags = new ArrayList<>();
    private Pattern searchPattern;
    private final Completer completer = new Completer();
    private Supplier<List<VbaSyntax.Completion>> extraSymbols = ArrayListSupplier.EMPTY;

    /** Empty supplier holder (avoids a lambda in a field initialiser order trap). */
    private static final class ArrayListSupplier {
        static final Supplier<List<VbaSyntax.Completion>> EMPTY = ArrayList::new;
    }

    public CodeEditor() {
        CodeFonts.register();
        family = CodeFonts.defaultFamily();
        setOpaque(false);
        setMargin(new Insets(6, 8, 6, 8));
        putClientProperty("caretWidth", 2);
        setDragEnabled(false);
        getStyledDocument();
        applyTheme(EditorTheme.forApp());
        applyFont();
        getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                changed(e.getOffset(), e.getLength());
            }

            @Override
            public void removeUpdate(DocumentEvent e) {
                changed(e.getOffset(), 0);
            }

            @Override
            public void changedUpdate(DocumentEvent e) {
                // attribute changes (our own colouring) need nothing
            }
        });
        getDocument().addUndoableEditListener(this::edit);
        addCaretListener((CaretEvent e) -> {
            SwingUtilities.invokeLater(() -> {
                updateBrackets();
                gutter.repaint();
                repaint();
            });
        });
        addMouseWheelListener(e -> {
            if (e.isControlDown()) {
                setFontSize(fontSize + (e.getWheelRotation() < 0 ? 1 : -1));
                e.consume();
            } else {
                e.getComponent().getParent().dispatchEvent(SwingUtilities.convertMouseEvent(this, e, e.getComponent().getParent()));
            }
        });
        addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                completer.hide();
                boundaryNext = true;
            }
        });
    }

    /** The row header (line numbers) to put into the scroll pane. */
    public JComponent getGutter() {
        return gutter;
    }

    public void setExtraSymbols(Supplier<List<VbaSyntax.Completion>> s) {
        extraSymbols = s == null ? ArrayListSupplier.EMPTY : s;
    }

    // ------------------------------------------------------------------
    // Look
    // ------------------------------------------------------------------

    public EditorTheme getEditorTheme() {
        return theme;
    }

    public void setEditorTheme(EditorTheme t) {
        applyTheme(t);
        highlightAll();
        gutter.repaint();
        if (getParent() instanceof javax.swing.JViewport) ((javax.swing.JViewport) getParent()).setBackground(t.bg);
        repaint();
    }

    private void applyTheme(EditorTheme t) {
        theme = t;
        for (int i = 0; i < attrs.length; i++) {
            SimpleAttributeSet a = new SimpleAttributeSet();
            StyleConstants.setForeground(a, t.kind[i]);
            if (i == VbaSyntax.COMMENT) StyleConstants.setItalic(a, true);
            attrs[i] = a;
        }
        setBackground(t.bg);
        setForeground(t.fg);
        setCaretColor(t.caret);
        setSelectionColor(t.selection);
        setSelectedTextColor(null);
        Style def = getStyledDocument().getStyle(StyleContext.DEFAULT_STYLE);
        StyleConstants.setForeground(def, t.fg);
    }

    public String getFontFamily() {
        return family;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setEditorFont(String fam, int size) {
        family = fam;
        fontSize = Math.max(8, Math.min(48, size));
        applyFont();
        highlightAll();
        gutter.revalidate();
        gutter.repaint();
    }

    public void setFontSize(int size) {
        int old = fontSize;
        int n = Math.max(8, Math.min(48, size));
        if (n == old) return;
        setEditorFont(family, n);
        firePropertyChange("editorFontSize", old, n);
    }

    private void applyFont() {
        Font f = new Font(family, Font.PLAIN, fontSize);
        setFont(f);
        Style def = getStyledDocument().getStyle(StyleContext.DEFAULT_STYLE);
        StyleConstants.setFontFamily(def, family);
        StyleConstants.setFontSize(def, fontSize);
        SimpleAttributeSet para = new SimpleAttributeSet();
        StyleConstants.setLineSpacing(para, 0.12f);
        getStyledDocument().setParagraphAttributes(0, Math.max(1, getDocument().getLength()), para, false);
    }

    private FontMetrics fm() {
        return getFontMetrics(getFont());
    }

    // ------------------------------------------------------------------
    // Layout: no wrapping
    // ------------------------------------------------------------------

    @Override
    public boolean getScrollableTracksViewportWidth() {
        return getParent() instanceof javax.swing.JViewport && getUI().getPreferredSize(this).width <= getParent().getWidth();
    }

    @Override
    public Dimension getPreferredSize() {
        Dimension d = super.getPreferredSize();
        return new Dimension(d.width + 40, d.height + 20);
    }

    // ------------------------------------------------------------------
    // Text access
    // ------------------------------------------------------------------

    public String getCode() {
        try {
            return getDocument().getText(0, getDocument().getLength());
        } catch (BadLocationException e) {
            return "";
        }
    }

    /** Replaces everything and clears the undo history. */
    public void setCode(String text) {
        applying = true;
        try {
            setText(text == null ? "" : text.replace("\r\n", "\n").replace('\r', '\n'));
        } finally {
            applying = false;
        }
        undo.discardAllEdits();
        current = null;
        errorLine = -1;
        setCaretPosition(0);
        highlightAll();
        gutter.revalidate();
        gutter.repaint();
    }

    private String[] lines() {
        return getCode().split("\n", -1);
    }

    public int getCaretLine() {
        return getDocument().getDefaultRootElement().getElementIndex(getCaretPosition()) + 1;
    }

    public int getCaretColumn() {
        Element root = getDocument().getDefaultRootElement();
        int l = root.getElementIndex(getCaretPosition());
        return getCaretPosition() - root.getElement(l).getStartOffset() + 1;
    }

    public int getLineCount() {
        return getDocument().getDefaultRootElement().getElementCount();
    }

    public void goToLineStart() {
        setCaretPosition(0);
        scrollRectToVisible(new Rectangle(0, 0, 1, 1));
    }

    public void goToLine(int line) {
        Element root = getDocument().getDefaultRootElement();
        int idx = Math.max(0, Math.min(root.getElementCount() - 1, line - 1));
        int start = root.getElement(idx).getStartOffset();
        int end = root.getElement(idx).getEndOffset() - 1;
        setCaretPosition(start);
        moveCaretPosition(Math.max(start, end));
        requestFocusInWindow();
        try {
            Rectangle r = modelToView2D(start).getBounds();
            r.height += 60;
            r.y -= 30;
            scrollRectToVisible(r);
        } catch (BadLocationException ignored) {
            // line vanished
        }
    }

    public void setErrorLine(int line) {
        errorLine = line;
        gutter.repaint();
        repaint();
    }

    public int getErrorLine() {
        return errorLine;
    }

    // ------------------------------------------------------------------
    // Highlighting
    // ------------------------------------------------------------------

    private void changed(int offset, int length) {
        if (highlighting) return;
        final int from = offset;
        final int to = offset + length;
        SwingUtilities.invokeLater(() -> {
            highlightRange(from, to);
            gutter.revalidate();
            gutter.repaint();
            repaint();
        });
        firePropertyChange("codeChanged", false, true);
    }

    public void highlightAll() {
        highlightRange(0, getDocument().getLength());
    }

    private void highlightRange(int from, int to) {
        StyledDocument doc = getStyledDocument();
        int len = doc.getLength();
        if (len == 0) return;
        Element root = doc.getDefaultRootElement();
        int l1 = root.getElementIndex(Math.max(0, Math.min(from, len)));
        int l2 = root.getElementIndex(Math.max(0, Math.min(to, len)));
        highlighting = true;
        try {
            for (int l = l1; l <= l2; l++) {
                Element e = root.getElement(l);
                int s = e.getStartOffset();
                int n = Math.min(e.getEndOffset(), len + 1) - 1 - s;
                if (n <= 0) continue;
                String text = doc.getText(s, n);
                doc.setCharacterAttributes(s, n, attrs[VbaSyntax.PLAIN], true);
                for (VbaSyntax.Span sp : VbaSyntax.spans(text)) {
                    if (sp.kind != VbaSyntax.PLAIN) doc.setCharacterAttributes(s + sp.start, sp.end - sp.start, attrs[sp.kind], true);
                }
            }
        } catch (BadLocationException ignored) {
            // document changed under us: the next change re-colours
        } finally {
            highlighting = false;
        }
    }

    // ------------------------------------------------------------------
    // Undo (typing is grouped into words / pauses)
    // ------------------------------------------------------------------

    private void edit(UndoableEditEvent e) {
        UndoableEdit ue = e.getEdit();
        if (ue instanceof AbstractDocument.DefaultDocumentEvent
                && ((AbstractDocument.DefaultDocumentEvent) ue).getType() == DocumentEvent.EventType.CHANGE) return; // colouring
        if (applying) return;
        long now = System.currentTimeMillis();
        if (explicitGroup && current != null) {
            current.addEdit(ue);
            return;
        }
        if (current != null && current.isInProgress() && !boundaryNext && now - lastEditTime < 1000) {
            current.addEdit(ue);
        } else {
            endGroup();
            current = new CompoundEdit();
            current.addEdit(ue);
            undo.addEdit(current);
        }
        boundaryNext = false;
        lastEditTime = now;
    }

    private void endGroup() {
        if (current != null && current.isInProgress()) current.end();
        current = null;
    }

    private void beginGroup() {
        endGroup();
        current = new CompoundEdit();
        undo.addEdit(current);
        explicitGroup = true;
    }

    private void finishGroup() {
        explicitGroup = false;
        endGroup();
        boundaryNext = true;
    }

    public void undoEdit() {
        endGroup();
        completer.hide();
        if (undo.canUndo()) undo.undo();
        highlightAll();
    }

    public void redoEdit() {
        endGroup();
        completer.hide();
        if (undo.canRedo()) undo.redo();
        highlightAll();
    }

    // ------------------------------------------------------------------
    // Keyboard
    // ------------------------------------------------------------------

    @Override
    protected void processKeyEvent(KeyEvent e) {
        if (e.getID() == KeyEvent.KEY_PRESSED && !e.isConsumed()) {
            int k = e.getKeyCode();
            boolean ctrl = e.isControlDown() || e.isMetaDown();
            boolean shift = e.isShiftDown();
            boolean alt = e.isAltDown();
            if (completer.isShowing()) {
                if (k == KeyEvent.VK_DOWN) { completer.move(1); e.consume(); return; }
                if (k == KeyEvent.VK_UP) { completer.move(-1); e.consume(); return; }
                if (k == KeyEvent.VK_PAGE_DOWN) { completer.move(8); e.consume(); return; }
                if (k == KeyEvent.VK_PAGE_UP) { completer.move(-8); e.consume(); return; }
                if (k == KeyEvent.VK_ENTER || k == KeyEvent.VK_TAB) { completer.accept(); e.consume(); return; }
                if (k == KeyEvent.VK_ESCAPE) { completer.hide(); e.consume(); return; }
            }
            if (ctrl && !shift && k == KeyEvent.VK_Z) { undoEdit(); e.consume(); return; }
            if (ctrl && (k == KeyEvent.VK_Y || (shift && k == KeyEvent.VK_Z))) { redoEdit(); e.consume(); return; }
            if (ctrl && k == KeyEvent.VK_SPACE) { completer.show(true); e.consume(); return; }
            if (ctrl && (k == KeyEvent.VK_SLASH || k == KeyEvent.VK_DIVIDE)) { toggleComment(); e.consume(); return; }
            if (ctrl && !shift && k == KeyEvent.VK_D) { duplicateLine(); e.consume(); return; }
            if (ctrl && shift && k == KeyEvent.VK_K) { deleteLine(); e.consume(); return; }
            if (ctrl && shift && k == KeyEvent.VK_F) { formatCode(); e.consume(); return; }
            if (ctrl && (k == KeyEvent.VK_0 || k == KeyEvent.VK_NUMPAD0)) { setFontSize(14); e.consume(); return; }
            if (ctrl && (k == KeyEvent.VK_EQUALS || k == KeyEvent.VK_PLUS || k == KeyEvent.VK_ADD)) { setFontSize(fontSize + 1); e.consume(); return; }
            if (ctrl && (k == KeyEvent.VK_MINUS || k == KeyEvent.VK_SUBTRACT)) { setFontSize(fontSize - 1); e.consume(); return; }
            if (alt && !ctrl && (k == KeyEvent.VK_UP || k == KeyEvent.VK_DOWN)) { moveLines(k == KeyEvent.VK_UP ? -1 : 1); e.consume(); return; }
            if (!ctrl && !alt && k == KeyEvent.VK_ENTER && isEditable()) { smartEnter(); e.consume(); return; }
            if (!ctrl && !alt && k == KeyEvent.VK_TAB && isEditable()) { tab(shift); e.consume(); return; }
            if (!ctrl && !alt && k == KeyEvent.VK_HOME) { smartHome(shift); e.consume(); return; }
            if (k == KeyEvent.VK_BACK_SPACE && !ctrl && !alt && isEditable() && smartBackspace()) { e.consume(); return; }
        }
        super.processKeyEvent(e);
        if (e.getID() == KeyEvent.KEY_TYPED && !e.isConsumed()) {
            char c = e.getKeyChar();
            if (Character.isLetterOrDigit(c) || c == '_' || c == '.') {
                SwingUtilities.invokeLater(() -> completer.show(false));
            } else {
                completer.hide();
            }
        }
        if (e.getID() == KeyEvent.KEY_PRESSED && (e.getKeyCode() == KeyEvent.VK_BACK_SPACE) && completer.isShowing()) {
            SwingUtilities.invokeLater(() -> completer.show(false));
        }
    }

    /** Typing ( or " adds the partner; typing the partner steps over it. */
    @Override
    public void replaceSelection(String content) {
        if (content != null && content.length() == 1 && isEditable() && getSelectionStart() == getSelectionEnd()) {
            char c = content.charAt(0);
            int pos = getCaretPosition();
            char next = charAt(pos);
            char prev = pos > 0 ? charAt(pos - 1) : 0;
            if ((c == ')' || c == '"') && next == c) {
                setCaretPosition(pos + 1);
                return;
            }
            if (c == '(' && (next == 0 || next == '\n' || next == ' ' || next == ')' || next == ',')) {
                super.replaceSelection("()");
                setCaretPosition(pos + 1);
                return;
            }
            if (c == '"' && !Character.isLetterOrDigit(prev) && prev != '"' && !insideString(pos) && (next == 0 || next == '\n' || next == ' ' || next == ')' || next == ',')) {
                super.replaceSelection("\"\"");
                setCaretPosition(pos + 1);
                return;
            }
            if (c == ' ' || c == '\n') boundaryNext = true;
        } else if (content != null && content.length() > 1) {
            boundaryNext = true;
        }
        super.replaceSelection(content);
    }

    private boolean insideString(int pos) {
        Element root = getDocument().getDefaultRootElement();
        int l = root.getElementIndex(pos);
        int s = root.getElement(l).getStartOffset();
        try {
            String before = getDocument().getText(s, pos - s);
            int q = 0;
            for (int i = 0; i < before.length(); i++) if (before.charAt(i) == '"') q++;
            return q % 2 == 1;
        } catch (BadLocationException e) {
            return false;
        }
    }

    private char charAt(int pos) {
        try {
            if (pos < 0 || pos >= getDocument().getLength()) return 0;
            return getDocument().getText(pos, 1).charAt(0);
        } catch (BadLocationException e) {
            return 0;
        }
    }

    private int lineStart(int lineIdx) {
        return getDocument().getDefaultRootElement().getElement(lineIdx).getStartOffset();
    }

    private int lineEnd(int lineIdx) { // offset of the newline (or end of text)
        Element root = getDocument().getDefaultRootElement();
        return Math.min(getDocument().getLength(), root.getElement(lineIdx).getEndOffset() - 1);
    }

    private void smartHome(boolean shift) {
        Element root = getDocument().getDefaultRootElement();
        int l = root.getElementIndex(getCaretPosition());
        int s = lineStart(l);
        String[] ls = lines();
        String text = ls[l];
        int ind = 0;
        while (ind < text.length() && text.charAt(ind) == ' ') ind++;
        int target = getCaretPosition() == s + ind ? s : s + ind;
        if (shift) moveCaretPosition(target);
        else setCaretPosition(target);
    }

    private boolean smartBackspace() {
        if (getSelectionStart() != getSelectionEnd()) return false;
        int pos = getCaretPosition();
        Element root = getDocument().getDefaultRootElement();
        int l = root.getElementIndex(pos);
        int s = lineStart(l);
        int col = pos - s;
        if (col <= 0) return false;
        try {
            String before = getDocument().getText(s, col);
            if (before.trim().isEmpty() && col % UNIT.length() == 0 || before.trim().isEmpty() && col > 1) {
                int remove = col % UNIT.length() == 0 ? UNIT.length() : col % UNIT.length();
                getDocument().remove(pos - remove, remove);
                return true;
            }
            // delete matching pair () or ""
            char a = before.charAt(col - 1);
            char b = charAt(pos);
            if ((a == '(' && b == ')') || (a == '"' && b == '"')) {
                getDocument().remove(pos - 1, 2);
                return true;
            }
        } catch (BadLocationException e) {
            return false;
        }
        return false;
    }

    private void tab(boolean shift) {
        int a = getSelectionStart();
        int b = getSelectionEnd();
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(a);
        int l2 = root.getElementIndex(b);
        if (a != b && root.getElement(l2).getStartOffset() == b && l2 > l1) l2--;
        if (!shift && a == b) {
            int pos = getCaretPosition();
            int col = pos - lineStart(l1);
            int n = UNIT.length() - (col % UNIT.length());
            super.replaceSelection(" ".repeat(n));
            return;
        }
        beginGroup();
        try {
            for (int l = l1; l <= l2; l++) {
                int s = lineStart(l);
                if (shift) {
                    int n = 0;
                    while (n < UNIT.length() && charAt(s + n) == ' ') n++;
                    if (n > 0) getDocument().remove(s, n);
                } else if (lineEnd(l) > s) {
                    getDocument().insertString(s, UNIT, null);
                }
            }
        } catch (BadLocationException ignored) {
            // nothing to do
        } finally {
            finishGroup();
        }
        selectLines(l1, l2);
    }

    private void selectLines(int l1, int l2) {
        setCaretPosition(lineStart(l1));
        moveCaretPosition(lineEnd(l2));
    }

    public void toggleComment() {
        if (!isEditable()) return;
        int a = getSelectionStart();
        int b = getSelectionEnd();
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(a);
        int l2 = root.getElementIndex(b);
        if (a != b && root.getElement(l2).getStartOffset() == b && l2 > l1) l2--;
        String[] ls = lines();
        boolean allCommented = true;
        boolean any = false;
        for (int l = l1; l <= l2; l++) {
            String t = ls[l].trim();
            if (t.isEmpty()) continue;
            any = true;
            if (!t.startsWith("'")) allCommented = false;
        }
        if (!any) return;
        beginGroup();
        try {
            for (int l = l1; l <= l2; l++) {
                String t = ls[l];
                if (t.trim().isEmpty()) continue;
                int ind = 0;
                while (ind < t.length() && t.charAt(ind) == ' ') ind++;
                int s = lineStart(l) + ind;
                if (allCommented) {
                    int n = t.startsWith("' ", ind) ? 2 : 1;
                    getDocument().remove(s, n);
                } else {
                    getDocument().insertString(s, "' ", null);
                }
            }
        } catch (BadLocationException ignored) {
            // nothing to do
        } finally {
            finishGroup();
        }
        selectLines(l1, l2);
    }

    private void duplicateLine() {
        if (!isEditable()) return;
        int a = getSelectionStart();
        int b = getSelectionEnd();
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(a);
        int l2 = root.getElementIndex(b);
        if (a != b && root.getElement(l2).getStartOffset() == b && l2 > l1) l2--;
        try {
            String text = getDocument().getText(lineStart(l1), lineEnd(l2) - lineStart(l1));
            int at = lineEnd(l2);
            beginGroup();
            getDocument().insertString(at, "\n" + text, null);
            finishGroup();
            selectLines(l2 + 1, l2 + 1 + (l2 - l1));
        } catch (BadLocationException ignored) {
            finishGroup();
        }
    }

    private void deleteLine() {
        if (!isEditable()) return;
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(getSelectionStart());
        int l2 = root.getElementIndex(getSelectionEnd());
        int s = lineStart(l1);
        int e = l2 + 1 < root.getElementCount() ? lineStart(l2 + 1) : getDocument().getLength();
        if (e == getDocument().getLength() && s > 0) s--; // take the previous newline for the last line
        try {
            getDocument().remove(s, e - s);
        } catch (BadLocationException ignored) {
            // nothing
        }
    }

    private void moveLines(int dir) {
        if (!isEditable()) return;
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(getSelectionStart());
        int l2 = root.getElementIndex(getSelectionEnd());
        if (getSelectionStart() != getSelectionEnd() && root.getElement(l2).getStartOffset() == getSelectionEnd() && l2 > l1) l2--;
        int n = root.getElementCount();
        if (dir < 0 && l1 == 0 || dir > 0 && l2 >= n - 1) return;
        String[] ls = lines();
        StringBuilder moved = new StringBuilder();
        for (int l = l1; l <= l2; l++) moved.append(l > l1 ? "\n" : "").append(ls[l]);
        int swap = dir < 0 ? l1 - 1 : l2 + 1;
        int from = dir < 0 ? lineStart(swap) : lineStart(l1);
        int to = dir < 0 ? lineEnd(l2) : lineEnd(swap);
        String text = dir < 0 ? moved + "\n" + ls[swap] : ls[swap] + "\n" + moved;
        try {
            beginGroup();
            getDocument().remove(from, to - from);
            getDocument().insertString(from, text, null);
        } catch (BadLocationException ignored) {
            // nothing
        } finally {
            finishGroup();
        }
        selectLines(l1 + dir, l2 + dir);
    }

    /** Enter: keeps the indent, outdents End/Next/Else lines, fixes keyword case and closes new blocks. */
    private void smartEnter() {
        try {
            beginGroup();
            if (getSelectionStart() != getSelectionEnd()) getDocument().remove(getSelectionStart(), getSelectionEnd() - getSelectionStart());
            Element root = getDocument().getDefaultRootElement();
            int pos = getCaretPosition();
            int l = root.getElementIndex(pos);
            String[] ls = lines();
            String line = ls[l];
            int s = lineStart(l);
            int col = pos - s;
            String before = line.substring(0, Math.min(col, line.length()));
            String after = line.substring(Math.min(col, line.length()));
            boolean inComment = false;
            for (VbaSyntax.Span sp : VbaSyntax.spans(before)) if (sp.kind == VbaSyntax.COMMENT) inComment = true;
            // 1. fix the finished line (indent + keyword case) - not inside a comment/string we are splitting
            String fixed = VbaSyntax.fixCase(before.trim());
            String[] probe = ls.clone();
            probe[l] = fixed + (after.trim().isEmpty() ? "" : after);
            int[] lv = VbaSyntax.levels(probe);
            String head = before.trim().isEmpty() ? "" : " ".repeat(lv[l] * UNIT.length()) + fixed;
            int lineEndOff = lineEnd(l);
            // replace the part of the line before the caret
            getDocument().remove(s, pos - s);
            getDocument().insertString(s, head, null);
            int caret = s + head.length();
            // 2. split the line
            getDocument().insertString(caret, "\n", null);
            String[] after2 = lines();
            int[] lv2 = VbaSyntax.levels(after2);
            int newLevel = l + 1 < lv2.length ? lv2[l + 1] : 0;
            // a Case line / opener makes the following line deeper
            String role = VbaSyntax.role(after2[l]);
            String ind = " ".repeat(newLevel * UNIT.length());
            if (!after.trim().isEmpty()) {
                // text moved to the new line: re-indent it too
                String t = VbaSyntax.fixCase(after.trim());
                int nl = caret + 1;
                getDocument().remove(nl, getDocument().getLength() >= nl ? lineEnd(l + 1) - nl : 0);
                getDocument().insertString(nl, ind + t, null);
                setCaretPosition(nl + ind.length());
            } else {
                int nl = caret + 1;
                getDocument().insertString(nl, ind, null);
                setCaretPosition(nl + ind.length());
                String closer = inComment ? null : VbaSyntax.closerFor(role);
                if (closer != null && !VbaSyntax.isClosed(after2, l) && newLevel >= 0) {
                    String closeIndent = " ".repeat(lv2[l] * UNIT.length());
                    int here = getCaretPosition();
                    getDocument().insertString(here, "\n" + closeIndent + closer, null);
                    setCaretPosition(here);
                }
            }
        } catch (BadLocationException ignored) {
            // nothing
        } finally {
            finishGroup();
        }
        scrollCaret();
    }

    private void scrollCaret() {
        try {
            Rectangle r = modelToView2D(getCaretPosition()).getBounds();
            r.height += 4;
            scrollRectToVisible(r);
        } catch (BadLocationException ignored) {
            // nothing
        }
    }

    /** Re-indents the whole module and fixes keyword capitalisation (one undo step). */
    public void formatCode() {
        if (!isEditable()) return;
        String old = getCode();
        String nw = VbaSyntax.format(old, UNIT);
        if (old.equals(nw)) return;
        int line = getCaretLine();
        beginGroup();
        try {
            getDocument().remove(0, getDocument().getLength());
            getDocument().insertString(0, nw, null);
        } catch (BadLocationException ignored) {
            // nothing
        } finally {
            finishGroup();
        }
        highlightAll();
        goToLine(line);
    }

    // ------------------------------------------------------------------
    // Search
    // ------------------------------------------------------------------

    /** Highlights every match and remembers the pattern for {@link #findNext}. @return match count (-1 for a bad regex) */
    public int markMatches(String text, boolean matchCase, boolean word, boolean regex) {
        Highlighter h = getHighlighter();
        for (Object t : matchTags) h.removeHighlight(t);
        matchTags.clear();
        searchPattern = null;
        if (text == null || text.isEmpty()) return 0;
        try {
            String p = regex ? text : Pattern.quote(text);
            if (word) p = "\\b(?:" + p + ")\\b";
            searchPattern = Pattern.compile(p, matchCase ? 0 : Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
        } catch (RuntimeException e) {
            return -1;
        }
        int n = 0;
        Matcher m = searchPattern.matcher(getCode());
        DefaultHighlighter.DefaultHighlightPainter painter = new DefaultHighlighter.DefaultHighlightPainter(theme.match);
        while (m.find() && n < 5000) {
            if (m.end() == m.start()) continue;
            try {
                matchTags.add(h.addHighlight(m.start(), m.end(), painter));
            } catch (BadLocationException ignored) {
                break;
            }
            n++;
        }
        return n;
    }

    public void clearMatches() {
        markMatches("", false, false, false);
    }

    /** Selects the next (or previous) match, wrapping around. */
    public boolean findNext(boolean forward) {
        if (searchPattern == null) return false;
        String code = getCode();
        Matcher m = searchPattern.matcher(code);
        int from = forward ? getSelectionEnd() : getSelectionStart();
        int s = -1;
        int e = -1;
        if (forward) {
            if (m.find(Math.min(from, code.length())) && m.end() > m.start()) {
                s = m.start();
                e = m.end();
            } else if (m.find(0) && m.end() > m.start()) {
                s = m.start();
                e = m.end();
            }
        } else {
            m.reset();
            while (m.find()) {
                if (m.end() == m.start()) continue;
                if (m.start() < from) {
                    s = m.start();
                    e = m.end();
                } else break;
            }
            if (s < 0) { // wrap: last match
                m.reset();
                while (m.find()) {
                    if (m.end() > m.start()) {
                        s = m.start();
                        e = m.end();
                    }
                }
            }
        }
        if (s < 0) return false;
        setCaretPosition(s);
        moveCaretPosition(e);
        scrollCaret();
        return true;
    }

    public boolean replaceCurrent(String repl) {
        if (searchPattern == null || !isEditable()) return false;
        int a = getSelectionStart();
        int b = getSelectionEnd();
        if (a != b) {
            Matcher m = searchPattern.matcher(getCode().substring(a, b));
            if (m.matches()) {
                try {
                    beginGroup();
                    getDocument().remove(a, b - a);
                    getDocument().insertString(a, m.replaceFirst(Matcher.quoteReplacement(repl)).isEmpty() ? "" : repl, null);
                } catch (BadLocationException ignored) {
                    // nothing
                } finally {
                    finishGroup();
                }
            }
        }
        return findNext(true);
    }

    public int replaceAll(String repl) {
        if (searchPattern == null || !isEditable()) return 0;
        String code = getCode();
        Matcher m = searchPattern.matcher(code);
        int n = 0;
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            if (m.end() == m.start()) continue;
            m.appendReplacement(sb, Matcher.quoteReplacement(repl));
            n++;
        }
        m.appendTail(sb);
        if (n > 0) {
            beginGroup();
            try {
                getDocument().remove(0, getDocument().getLength());
                getDocument().insertString(0, sb.toString(), null);
            } catch (BadLocationException ignored) {
                // nothing
            } finally {
                finishGroup();
            }
            highlightAll();
        }
        return n;
    }

    // ------------------------------------------------------------------
    // Brackets
    // ------------------------------------------------------------------

    private void updateBrackets() {
        Highlighter h = getHighlighter();
        if (bracketA != null) h.removeHighlight(bracketA);
        if (bracketB != null) h.removeHighlight(bracketB);
        bracketA = null;
        bracketB = null;
        int pos = getCaretPosition();
        int at = -1;
        char c = 0;
        if (charAt(pos) == '(' || charAt(pos) == ')') {
            at = pos;
            c = charAt(pos);
        } else if (pos > 0 && (charAt(pos - 1) == '(' || charAt(pos - 1) == ')')) {
            at = pos - 1;
            c = charAt(pos - 1);
        }
        if (at < 0) return;
        Element root = getDocument().getDefaultRootElement();
        int l = root.getElementIndex(at);
        int s = lineStart(l);
        int e = lineEnd(l);
        try {
            String line = getDocument().getText(s, e - s);
            int idx = at - s;
            int depth = 0;
            int match = -1;
            if (c == '(') {
                for (int i = idx; i < line.length(); i++) {
                    if (line.charAt(i) == '(') depth++;
                    else if (line.charAt(i) == ')' && --depth == 0) {
                        match = i;
                        break;
                    }
                }
            } else {
                for (int i = idx; i >= 0; i--) {
                    if (line.charAt(i) == ')') depth++;
                    else if (line.charAt(i) == '(' && --depth == 0) {
                        match = i;
                        break;
                    }
                }
            }
            Color col = new Color(theme.bracket.getRed(), theme.bracket.getGreen(), theme.bracket.getBlue(), match >= 0 ? 110 : 0);
            DefaultHighlighter.DefaultHighlightPainter p = new DefaultHighlighter.DefaultHighlightPainter(
                    match >= 0 ? col : new Color(0xE53935));
            bracketA = h.addHighlight(at, at + 1, p);
            if (match >= 0) bracketB = h.addHighlight(s + match, s + match + 1, p);
        } catch (BadLocationException ignored) {
            // nothing
        }
    }

    // ------------------------------------------------------------------
    // Painting
    // ------------------------------------------------------------------

    @Override
    protected void paintComponent(Graphics g0) {
        Graphics2D g = (Graphics2D) g0.create();
        g.setColor(theme.bg);
        g.fillRect(0, 0, getWidth(), getHeight());
        try {
            Element root = getDocument().getDefaultRootElement();
            Rectangle clip = g.getClipBounds() == null ? new Rectangle(0, 0, getWidth(), getHeight()) : g.getClipBounds();
            if (getSelectionStart() == getSelectionEnd()) {
                int l = root.getElementIndex(getCaretPosition());
                Rectangle r = modelToView2D(root.getElement(l).getStartOffset()).getBounds();
                g.setColor(theme.currentLine);
                g.fillRect(0, r.y, getWidth(), r.height);
            }
            if (errorLine > 0 && errorLine <= root.getElementCount()) {
                Rectangle r = modelToView2D(root.getElement(errorLine - 1).getStartOffset()).getBounds();
                g.setColor(theme.errorLine);
                g.fillRect(0, r.y, getWidth(), r.height);
            }
        } catch (BadLocationException ignored) {
            // nothing
        }
        g.dispose();
        super.paintComponent(g0);
        Graphics2D g2 = (Graphics2D) g0.create();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        try {
            paintGuides(g2);
            paintSquiggle(g2);
        } catch (BadLocationException ignored) {
            // nothing
        } finally {
            g2.dispose();
        }
    }

    private void paintGuides(Graphics2D g) throws BadLocationException {
        Rectangle vis = getVisibleRect();
        Element root = getDocument().getDefaultRootElement();
        int l1 = root.getElementIndex(viewToModel2D(new Point(0, vis.y)));
        int l2 = root.getElementIndex(viewToModel2D(new Point(0, vis.y + vis.height)));
        FontMetrics fm = fm();
        int cw = fm.charWidth('M');
        int left = getInsets().left;
        g.setColor(theme.guide);
        for (int l = l1; l <= l2 && l < root.getElementCount(); l++) {
            Element el = root.getElement(l);
            int s = el.getStartOffset();
            int n = el.getEndOffset() - 1 - s;
            if (n <= 0) continue;
            String t = getDocument().getText(s, n);
            int ind = 0;
            while (ind < t.length() && t.charAt(ind) == ' ') ind++;
            if (ind < UNIT.length()) continue;
            Rectangle r = modelToView2D(s).getBounds();
            for (int c = UNIT.length(); c <= ind; c += UNIT.length()) {
                int x = left + (c - UNIT.length()) * cw;
                g.drawLine(x, r.y, x, r.y + r.height);
            }
        }
    }

    private void paintSquiggle(Graphics2D g) throws BadLocationException {
        Element root = getDocument().getDefaultRootElement();
        if (errorLine <= 0 || errorLine > root.getElementCount()) return;
        Element el = root.getElement(errorLine - 1);
        int s = el.getStartOffset();
        int e = el.getEndOffset() - 1;
        String t = getDocument().getText(s, Math.max(0, e - s));
        int ind = 0;
        while (ind < t.length() && t.charAt(ind) == ' ') ind++;
        if (ind >= t.length()) return;
        Rectangle a = modelToView2D(s + ind).getBounds();
        Rectangle b = modelToView2D(e).getBounds();
        int y = a.y + a.height - 2;
        g.setColor(new Color(0xE53935));
        g.setStroke(new BasicStroke(1.4f));
        int x = a.x;
        int up = 0;
        java.awt.geom.Path2D.Float path = new java.awt.geom.Path2D.Float();
        path.moveTo(x, y);
        while (x < b.x) {
            x += 3;
            up ^= 1;
            path.lineTo(x, y - (up == 1 ? 2 : 0));
        }
        g.draw(path);
    }

    // ------------------------------------------------------------------
    // Gutter
    // ------------------------------------------------------------------

    private final class Gutter extends JComponent {
        Gutter() {
            setOpaque(true);
            addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    try {
                        int off = viewToModel2D(new Point(0, e.getY()));
                        int l = getDocument().getDefaultRootElement().getElementIndex(off);
                        selectLines(l, l);
                        requestFocusInWindow();
                    } catch (RuntimeException ignored) {
                        // nothing
                    }
                }
            });
        }

        @Override
        public Dimension getPreferredSize() {
            int digits = Math.max(3, String.valueOf(Math.max(1, getLineCount())).length());
            FontMetrics m = getFontMetrics(gutterFont());
            return new Dimension(m.charWidth('0') * digits + 26, 10);
        }

        private Font gutterFont() {
            return new Font(family, Font.PLAIN, Math.max(8, fontSize - 2));
        }

        @Override
        protected void paintComponent(Graphics g0) {
            Graphics2D g = (Graphics2D) g0.create();
            g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g.setColor(theme.gutterBg);
            g.fillRect(0, 0, getWidth(), getHeight());
            g.setColor(theme.border);
            g.drawLine(getWidth() - 1, 0, getWidth() - 1, getHeight());
            g.setFont(gutterFont());
            FontMetrics fm = g.getFontMetrics();
            Element root = getDocument().getDefaultRootElement();
            Rectangle clip = g.getClipBounds();
            try {
                int l1 = root.getElementIndex(viewToModel2D(new Point(0, clip.y)));
                int l2 = root.getElementIndex(viewToModel2D(new Point(0, clip.y + clip.height)));
                int cur = root.getElementIndex(getCaretPosition());
                for (int l = l1; l <= l2 && l < root.getElementCount(); l++) {
                    Rectangle r = modelToView2D(root.getElement(l).getStartOffset()).getBounds();
                    String num = String.valueOf(l + 1);
                    boolean isErr = l + 1 == errorLine;
                    if (isErr) {
                        g.setColor(new Color(0xE53935));
                        g.fillOval(4, r.y + (r.height - 8) / 2, 8, 8);
                    }
                    g.setColor(l == cur ? theme.gutterCurrent : theme.gutterFg);
                    int x = getWidth() - 8 - fm.stringWidth(num);
                    int by = r.y + (r.height - fm.getHeight()) / 2 + fm.getAscent();
                    g.drawString(num, x, by);
                }
            } catch (BadLocationException ignored) {
                // document changed while painting
            } finally {
                g.dispose();
            }
        }
    }

    // ------------------------------------------------------------------
    // Completion popup
    // ------------------------------------------------------------------

    private final class Completer {
        private JWindow win;
        private final DefaultListModel<VbaSyntax.Completion> model = new DefaultListModel<>();
        private JList<VbaSyntax.Completion> list;
        private int wordStart;

        boolean isShowing() {
            return win != null && win.isVisible();
        }

        void hide() {
            if (win != null) win.setVisible(false);
        }

        void move(int d) {
            int i = Math.max(0, Math.min(model.size() - 1, list.getSelectedIndex() + d));
            list.setSelectedIndex(i);
            list.ensureIndexIsVisible(i);
        }

        void accept() {
            VbaSyntax.Completion c = list.getSelectedValue();
            hide();
            if (c == null) return;
            try {
                beginGroup();
                int caret = getCaretPosition();
                getDocument().remove(wordStart, caret - wordStart);
                getDocument().insertString(wordStart, c.text, null);
            } catch (BadLocationException ignored) {
                // nothing
            } finally {
                finishGroup();
            }
        }

        void show(boolean manual) {
            if (!isEditable() || !isShowing() && !isDisplayable()) return;
            int caret = getCaretPosition();
            String code;
            try {
                Element root = getDocument().getDefaultRootElement();
                int l = root.getElementIndex(caret);
                int s = root.getElement(l).getStartOffset();
                code = getDocument().getText(s, caret - s);
                int i = code.length();
                while (i > 0 && (Character.isLetterOrDigit(code.charAt(i - 1)) || code.charAt(i - 1) == '_')) i--;
                String prefix = code.substring(i);
                wordStart = s + i;
                // inside a comment or string: no popup
                for (VbaSyntax.Span sp : VbaSyntax.spans(code)) {
                    if ((sp.kind == VbaSyntax.COMMENT || sp.kind == VbaSyntax.STRING) && sp.end >= code.length() && sp.start < code.length()) {
                        hide();
                        return;
                    }
                }
                boolean member = i > 0 && code.charAt(i - 1) == '.';
                if (!manual && !member && prefix.length() < 2) {
                    hide();
                    return;
                }
                if (!manual && member && prefix.length() == 0 && code.length() > 0 && code.charAt(code.length() - 1) != '.') {
                    hide();
                    return;
                }
                List<VbaSyntax.Completion> pool;
                if (member) {
                    String before = identifierBeforeDot(code, i - 1);
                    pool = VbaSyntax.membersFor(before);
                } else {
                    pool = new ArrayList<>(VbaSyntax.globalCompletions());
                    pool.addAll(userSymbols());
                    pool.addAll(extraSymbols.get());
                }
                String low = prefix.toLowerCase(Locale.ROOT);
                List<VbaSyntax.Completion> hits = new ArrayList<>();
                Set<String> seen = new HashSet<>();
                for (VbaSyntax.Completion c : pool) {
                    if (c.text.toLowerCase(Locale.ROOT).startsWith(low) && seen.add(c.text.toLowerCase(Locale.ROOT))) hits.add(c);
                }
                hits.sort((a, b) -> {
                    int ra = rank(a);
                    int rb = rank(b);
                    return ra != rb ? Integer.compare(ra, rb) : a.text.compareToIgnoreCase(b.text);
                });
                if (hits.isEmpty() || (hits.size() == 1 && hits.get(0).text.equalsIgnoreCase(prefix))) {
                    hide();
                    return;
                }
                display(hits, caret);
            } catch (BadLocationException | RuntimeException e) {
                hide();
            }
        }

        private int rank(VbaSyntax.Completion c) {
            switch (c.kind) {
                case "V": case "P": return 0;
                case "M": return 1;
                case "K": return 2;
                default: return 3;
            }
        }

        private String identifierBeforeDot(String code, int dot) {
            int e = dot;
            // skip a balanced (...) group
            if (e > 0 && code.charAt(e - 1) == ')') {
                int depth = 0;
                int i = e - 1;
                for (; i >= 0; i--) {
                    if (code.charAt(i) == ')') depth++;
                    else if (code.charAt(i) == '(' && --depth == 0) break;
                }
                e = Math.max(0, i);
            }
            int s = e;
            while (s > 0 && (Character.isLetterOrDigit(code.charAt(s - 1)) || code.charAt(s - 1) == '_')) s--;
            return code.substring(s, e);
        }

        private void display(List<VbaSyntax.Completion> hits, int caret) throws BadLocationException {
            if (win == null) {
                Window owner = SwingUtilities.getWindowAncestor(CodeEditor.this);
                win = new JWindow(owner);
                win.setFocusableWindowState(false);
                list = new JList<>(model);
                list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                list.setFocusable(false);
                list.setCellRenderer(new DefaultListCellRenderer() {
                    @Override
                    public java.awt.Component getListCellRendererComponent(JList<?> l, Object v, int i, boolean sel, boolean foc) {
                        VbaSyntax.Completion c = (VbaSyntax.Completion) v;
                        javax.swing.JLabel lab = (javax.swing.JLabel) super.getListCellRendererComponent(l, v, i, sel, false);
                        lab.setText(badge(c.kind) + "  " + c.text + "    " + c.detail);
                        lab.setFont(new Font(family, Font.PLAIN, Math.max(10, fontSize - 1)));
                        lab.setForeground(sel ? Color.WHITE : theme.fg);
                        lab.setBackground(sel ? theme.selection.darker() : theme.gutterBg);
                        lab.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 6, 1, 10));
                        return lab;
                    }
                });
                list.addMouseListener(new MouseAdapter() {
                    @Override
                    public void mouseClicked(MouseEvent e) {
                        if (e.getClickCount() >= 2) accept();
                    }
                });
                JScrollPane sp = new JScrollPane(list);
                sp.setBorder(javax.swing.BorderFactory.createLineBorder(theme.border));
                win.setContentPane(sp);
            }
            ((JScrollPane) win.getContentPane()).setBorder(javax.swing.BorderFactory.createLineBorder(theme.border));
            list.setBackground(theme.gutterBg);
            model.clear();
            for (VbaSyntax.Completion c : hits) model.addElement(c);
            list.setSelectedIndex(0);
            FontMetrics fm = getFontMetrics(new Font(family, Font.PLAIN, Math.max(10, fontSize - 1)));
            int rows = Math.min(10, hits.size());
            int widest = 0;
            for (VbaSyntax.Completion c : hits) widest = Math.max(widest, fm.stringWidth("W  " + c.text + "    " + c.detail));
            int w = Math.max(240, Math.min(560, widest + 44));
            int h = rows * (fm.getHeight() + 2) + 4;
            win.setSize(w, h);
            Rectangle r = modelToView2D(wordStart).getBounds();
            Point p = new Point(r.x, r.y + r.height);
            SwingUtilities.convertPointToScreen(p, CodeEditor.this);
            win.setLocation(p);
            win.setVisible(true);
        }

        private String badge(String kind) {
            switch (kind) {
                case "K": return "■";
                case "F": return "ƒ";
                case "T": return "T";
                case "O": return "◆";
                case "C": return "π";
                case "M": return "●";
                case "P": return "▶";
                default: return "·";
            }
        }
    }

    /** Sub/Function names and declared variables in this module, for the completion list. */
    private List<VbaSyntax.Completion> userSymbols() {
        List<VbaSyntax.Completion> out = new ArrayList<>();
        String code = getCode();
        Matcher m = Pattern.compile("(?im)^\\s*(?:(?:public|private|friend|static)\\s+)?(sub|function)\\s+([A-Za-z_][A-Za-z0-9_]*)").matcher(code);
        while (m.find()) out.add(new VbaSyntax.Completion(m.group(2), "P", m.group(1).toLowerCase(Locale.ROOT)));
        Matcher d = Pattern.compile("(?im)^\\s*(?:dim|const|static|private|public)\\s+([A-Za-z_][A-Za-z0-9_]*)").matcher(code);
        while (d.find()) out.add(new VbaSyntax.Completion(d.group(1), "V", "variable"));
        Matcher f = Pattern.compile("(?im)^\\s*for(?:\\s+each)?\\s+([A-Za-z_][A-Za-z0-9_]*)").matcher(code);
        while (f.find()) out.add(new VbaSyntax.Completion(f.group(1), "V", "variable"));
        return out;
    }
}
