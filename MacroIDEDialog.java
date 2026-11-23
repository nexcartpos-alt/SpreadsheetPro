import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Insets;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.DefaultListCellRenderer;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JRadioButtonMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.ListSelectionModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingUtilities;
import javax.swing.Timer;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.Element;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;

/**
 * Macro Studio: a code editor with syntax colours, several modules saved inside the workbook, a procedure
 * picker, an Immediate window, find and replace, import/export of .bas files, fonts and themes.
 */
public class MacroIDEDialog extends JDialog {

    /** What the studio needs from the application. */
    public interface Host {
        SpreadsheetModel activeModel();

        MacroEngine.Workbook workbook();

        /** The live list of modules stored in the workbook. */
        List<MacroModule> modules();

        /** Cells or modules were changed. */
        void changed();

        String workbookName();
    }

    private static final Color GREEN = new Color(0x107C41);
    private static final String STARTER = "Sub Main()\n    ' Write your macro here, then press F5 to run it\n"
            + "    Range(\"A1\").Value = \"Hello from Spreadsheet Pro\"\nEnd Sub\n";

    private final Host host;
    private final CodeEditor editor = new CodeEditor();
    private final JScrollPane editorScroll = new JScrollPane(editor);
    private final DefaultListModel<MacroModule> moduleModel = new DefaultListModel<>();
    private final JList<MacroModule> moduleList = new JList<>(moduleModel);
    private final JTextPane output = new JTextPane();
    private final JTextField immediate = new JTextField();
    private final List<String> immediateHistory = new ArrayList<>();
    private int historyPos;
    private final JComboBox<ProcRef> procCombo = new JComboBox<>();
    private final JLabel posLabel = new JLabel(" ");
    private final JLabel infoLabel = new JLabel(" ");
    private final JLabel fontLabel = new JLabel(" ");
    private final JPanel findBar = new JPanel(new BorderLayout());
    private final JTextField findField = new JTextField(18);
    private final JTextField replField = new JTextField(14);
    private final JCheckBox caseBox = new JCheckBox("Aa");
    private final JCheckBox wordBox = new JCheckBox("Word");
    private final JCheckBox regexBox = new JCheckBox(".*");
    private final JLabel findCount = new JLabel(" ");
    private final JPanel replRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 2));
    private final JButton runBtn = new JButton("Run  F5");
    private final Timer syncTimer;
    private MacroModule current;
    private boolean updatingCombo;
    private boolean switching;
    private final JPanel left = new JPanel(new BorderLayout());
    private final JPanel bottom = new JPanel(new BorderLayout());
    private final JPanel outHead = new JPanel(new BorderLayout());
    private final JPanel immRow = new JPanel(new BorderLayout(6, 0));
    private final JLabel immLabel = new JLabel("Immediate ›");

    /** A procedure shown in the drop-down. */
    private static final class ProcRef {
        final int module;
        final String moduleName;
        final MacroEngine.ProcInfo info;

        ProcRef(int module, String moduleName, MacroEngine.ProcInfo info) {
            this.module = module;
            this.moduleName = moduleName;
            this.info = info;
        }

        @Override
        public String toString() {
            return (info.function ? "ƒ  " : "▶  ") + info.name + "  —  " + moduleName;
        }
    }

    public MacroIDEDialog(JFrame owner, Host host) {
        super(owner, "Macro Studio", false);
        this.host = host;
        setDefaultCloseOperation(HIDE_ON_CLOSE);

        syncTimer = new Timer(450, e -> {
            flush();
            refreshProcedures();
        });
        syncTimer.setRepeats(false);

        buildEditor();
        buildLeft();
        buildBottom();
        buildFindBar();

        JPanel top = new JPanel(new BorderLayout());
        top.add(buildToolbar(), BorderLayout.NORTH);
        top.add(findBar, BorderLayout.SOUTH);

        JSplitPane centre = new JSplitPane(JSplitPane.VERTICAL_SPLIT, editorScroll, bottom);
        centre.setResizeWeight(0.74);
        centre.setBorder(null);
        JSplitPane main = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, left, centre);
        main.setDividerLocation(190);
        main.setBorder(null);

        JPanel status = new JPanel(new FlowLayout(FlowLayout.LEFT, 18, 3));
        status.setBorder(BorderFactory.createMatteBorder(1, 0, 0, 0, new Color(0x80808040, true)));
        status.add(posLabel);
        status.add(infoLabel);
        status.add(fontLabel);

        add(top, BorderLayout.NORTH);
        add(main, BorderLayout.CENTER);
        add(status, BorderLayout.SOUTH);

        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                flush();
            }
        });
        installKeys();

        // restore preferences
        String fam = AppSettings.get("macro.font", "");
        int size = 14;
        try {
            size = Integer.parseInt(AppSettings.get("macro.size", "14"));
        } catch (NumberFormatException ignored) {
            // default
        }
        if (!fam.isEmpty() && CodeFonts.monospaced().contains(fam)) editor.setEditorFont(fam, size);
        else editor.setEditorFont(CodeFonts.defaultFamily(), size);
        String tn = AppSettings.get("macro.theme", "");
        EditorTheme et = EditorTheme.byName(tn);
        applyTheme(et != null ? et : EditorTheme.forApp());
        editor.addPropertyChangeListener("editorFontSize", e -> {
            AppSettings.put("macro.size", String.valueOf(editor.getFontSize()));
            updateFontLabel();
        });
        editor.addPropertyChangeListener("codeChanged", e -> {
            if (!switching) {
                editor.setErrorLine(-1);
                syncTimer.restart();
            }
        });
        editor.addCaretListener(e -> updatePos());
        editor.setExtraSymbols(this::otherModuleSymbols);

        loadModules();
        setSize(1120, 720);
        setLocationRelativeTo(owner);
        updateFontLabel();
    }

    // ------------------------------------------------------------------
    // Construction
    // ------------------------------------------------------------------

    private void buildEditor() {
        editorScroll.setRowHeaderView(editor.getGutter());
        editorScroll.setBorder(null);
        editorScroll.getVerticalScrollBar().setUnitIncrement(18);
    }

    private void buildLeft() {
        JLabel head = new JLabel("  Modules");
        head.setFont(head.getFont().deriveFont(Font.BOLD, 12f));
        head.setBorder(BorderFactory.createEmptyBorder(6, 4, 6, 4));
        moduleList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        moduleList.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> l, Object v, int i, boolean sel, boolean foc) {
                JLabel lab = (JLabel) super.getListCellRendererComponent(l, v, i, sel, false);
                MacroModule m = (MacroModule) v;
                int n = MacroEngine.procedures(m == current ? editor.getCode() : m.code).size();
                lab.setText("<html>▤&nbsp; <b>" + m.name + "</b> <span style='color:gray'>&nbsp;" + n + (n == 1 ? " proc" : " procs") + "</span></html>");
                lab.setBorder(BorderFactory.createEmptyBorder(5, 8, 5, 8));
                return lab;
            }
        });
        moduleList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting() && moduleList.getSelectedValue() != null && moduleList.getSelectedValue() != current) {
                switchTo(moduleList.getSelectedValue());
            }
        });
        moduleList.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) renameModule();
            }
        });
        JPanel btns = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 4));
        btns.add(small("+", "New module", this::newModule));
        btns.add(small("✎", "Rename module", this::renameModule));
        btns.add(small("✖", "Delete module", this::deleteModule));
        left.add(head, BorderLayout.NORTH);
        left.add(new JScrollPane(moduleList), BorderLayout.CENTER);
        left.add(btns, BorderLayout.SOUTH);
        left.setMinimumSize(new Dimension(150, 100));
    }

    private JButton small(String text, String tip, Runnable r) {
        JButton b = new JButton(text);
        b.setToolTipText(tip);
        b.setMargin(new Insets(1, 6, 1, 6));
        b.setFocusable(false);
        b.addActionListener(e -> r.run());
        return b;
    }

    private void buildBottom() {
        output.setEditable(false);
        output.setMargin(new Insets(6, 8, 6, 8));
        JLabel title = new JLabel("  Output");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 12f));
        JButton clear = small("Clear", "Clear the output", () -> output.setText(""));
        outHead.add(title, BorderLayout.WEST);
        JPanel cw = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 2));
        cw.setOpaque(false);
        cw.add(clear);
        outHead.add(cw, BorderLayout.EAST);
        outHead.setBorder(BorderFactory.createEmptyBorder(2, 0, 2, 0));

        immediate.setToolTipText("Type a statement and press Enter, or start with ? to print a value, e.g.  ? 2 + 2");
        immediate.addActionListener(e -> runImmediate());
        immediate.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if (immediateHistory.isEmpty()) return;
                if (e.getKeyCode() == KeyEvent.VK_UP) {
                    historyPos = Math.max(0, historyPos - 1);
                    immediate.setText(immediateHistory.get(historyPos));
                    e.consume();
                } else if (e.getKeyCode() == KeyEvent.VK_DOWN) {
                    historyPos = Math.min(immediateHistory.size(), historyPos + 1);
                    immediate.setText(historyPos < immediateHistory.size() ? immediateHistory.get(historyPos) : "");
                    e.consume();
                }
            }
        });
        immRow.add(immLabel, BorderLayout.WEST);
        immRow.add(immediate, BorderLayout.CENTER);
        immRow.setBorder(BorderFactory.createEmptyBorder(4, 8, 6, 8));

        output.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                int off = output.viewToModel2D(e.getPoint());
                Element el = output.getStyledDocument().getCharacterElement(off);
                Object g = el.getAttributes().getAttribute("goto");
                if (g instanceof int[]) jumpTo(((int[]) g)[0], ((int[]) g)[1]);
            }
        });
        output.addMouseMotionListener(new MouseAdapter() {
            @Override
            public void mouseMoved(MouseEvent e) {
                int off = output.viewToModel2D(e.getPoint());
                Object g = output.getStyledDocument().getCharacterElement(off).getAttributes().getAttribute("goto");
                output.setCursor(java.awt.Cursor.getPredefinedCursor(g instanceof int[] ? java.awt.Cursor.HAND_CURSOR : java.awt.Cursor.TEXT_CURSOR));
            }
        });
        bottom.add(outHead, BorderLayout.NORTH);
        bottom.add(new JScrollPane(output), BorderLayout.CENTER);
        bottom.add(immRow, BorderLayout.SOUTH);
        bottom.setMinimumSize(new Dimension(100, 90));
    }

    private JComponent buildToolbar() {
        JPanel bar = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 5));
        bar.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(0x80808040, true)));
        runBtn.setIcon(Icons.get("play", 16, Color.WHITE));
        runBtn.setBackground(GREEN);
        runBtn.setForeground(Color.WHITE);
        runBtn.setToolTipText("Run the procedure under the cursor (or the one chosen in the list)  -  F5");
        runBtn.setFocusable(false);
        runBtn.addActionListener(e -> runCurrent());
        bar.add(runBtn);
        procCombo.setPreferredSize(new Dimension(250, 28));
        procCombo.setToolTipText("Procedures in all modules - pick one to jump to it");
        procCombo.addActionListener(e -> {
            if (updatingCombo) return;
            ProcRef p = (ProcRef) procCombo.getSelectedItem();
            if (p != null) jumpTo(p.module, p.info.line);
        });
        bar.add(procCombo);
        bar.add(Box.createHorizontalStrut(6));
        bar.add(tool("Format", "indent", "Tidy indentation and keyword capitals  -  Ctrl+Shift+F", editor::formatCode));
        bar.add(tool("Find", "find", "Find and replace in this module  -  Ctrl+F / Ctrl+H", () -> showFind(false)));
        bar.add(Box.createHorizontalStrut(6));
        bar.add(tool("Import", "open", "Load a .bas file as a new module", this::importBas));
        bar.add(tool("Export", "save", "Save this module as a .bas file", this::exportBas));
        bar.add(Box.createHorizontalStrut(6));
        bar.add(tool("Samples", "macro", "Add an example macro as a new module", null, this::samplesMenu));
        bar.add(tool("Font", "pen", "Choose the editor font and size", this::fontDialog));
        bar.add(tool("Theme", "palette", "Editor colours", null, this::themeMenu));
        bar.add(tool("A−", null, "Smaller text  -  Ctrl+minus", () -> editor.setFontSize(editor.getFontSize() - 1)));
        bar.add(tool("A+", null, "Larger text  -  Ctrl+plus  (or Ctrl + mouse wheel)", () -> editor.setFontSize(editor.getFontSize() + 1)));
        bar.add(tool("Help", "help", "Language reference and shortcuts", this::help));
        return bar;
    }

    private JButton tool(String text, String icon, String tip, Runnable r) {
        return tool(text, icon, tip, r, null);
    }

    private JButton tool(String text, String icon, String tip, Runnable r, java.util.function.Consumer<JButton> popup) {
        JButton b = new JButton(text);
        if (icon != null) b.setIcon(Icons.get(icon, 16, Theme.fg));
        b.setToolTipText(tip);
        b.setFocusable(false);
        b.setMargin(new Insets(3, 8, 3, 8));
        if (popup != null) b.addActionListener(e -> popup.accept(b));
        else b.addActionListener(e -> r.run());
        return b;
    }

    private void buildFindBar() {
        JPanel row1 = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 2));
        row1.add(new JLabel("Find"));
        row1.add(findField);
        row1.add(caseBox);
        row1.add(wordBox);
        row1.add(regexBox);
        row1.add(small("▲", "Previous match  -  Shift+Enter", () -> editor.findNext(false)));
        row1.add(small("▼", "Next match  -  Enter", () -> editor.findNext(true)));
        row1.add(findCount);
        row1.add(small("✕", "Close  -  Esc", () -> hideFind()));
        replRow.add(new JLabel("Replace"));
        replRow.add(replField);
        replRow.add(small("Replace", "Replace this match", () -> {
            editor.replaceCurrent(replField.getText());
            updateMatches();
        }));
        replRow.add(small("Replace all", "Replace every match", () -> {
            int n = editor.replaceAll(replField.getText());
            updateMatches();
            findCount.setText(n + " replaced");
        }));
        JPanel rows = new JPanel(new BorderLayout());
        rows.add(row1, BorderLayout.NORTH);
        rows.add(replRow, BorderLayout.SOUTH);
        findBar.add(rows, BorderLayout.CENTER);
        findBar.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(0x80808040, true)));
        findBar.setVisible(false);
        javax.swing.event.DocumentListener dl = new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e) { updateMatches(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { updateMatches(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { updateMatches(); }
        };
        findField.getDocument().addDocumentListener(dl);
        caseBox.addActionListener(e -> updateMatches());
        wordBox.addActionListener(e -> updateMatches());
        regexBox.addActionListener(e -> updateMatches());
        findField.addActionListener(e -> editor.findNext(true));
        findField.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if (e.getKeyCode() == KeyEvent.VK_ENTER && e.isShiftDown()) {
                    editor.findNext(false);
                    e.consume();
                } else if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
                    hideFind();
                }
            }
        });
        replField.addActionListener(e -> {
            editor.replaceCurrent(replField.getText());
            updateMatches();
        });
    }

    private void installKeys() {
        JComponent root = getRootPane();
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0), this::runCurrent);
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_F, java.awt.event.InputEvent.CTRL_DOWN_MASK), () -> showFind(false));
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_H, java.awt.event.InputEvent.CTRL_DOWN_MASK), () -> showFind(true));
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_G, java.awt.event.InputEvent.CTRL_DOWN_MASK), this::gotoLineDialog);
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_F3, 0), () -> editor.findNext(true));
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_F3, java.awt.event.InputEvent.SHIFT_DOWN_MASK), () -> editor.findNext(false));
        bind(root, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), () -> {
            if (findBar.isVisible()) hideFind();
        });
    }

    private void bind(JComponent c, KeyStroke ks, Runnable r) {
        String key = "macro." + ks;
        c.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(ks, key);
        c.getActionMap().put(key, new AbstractAction() {
            @Override
            public void actionPerformed(java.awt.event.ActionEvent e) {
                r.run();
            }
        });
    }

    // ------------------------------------------------------------------
    // Look
    // ------------------------------------------------------------------

    private void applyTheme(EditorTheme t) {
        editor.setEditorTheme(t);
        editorScroll.getViewport().setBackground(t.bg);
        editorScroll.getRowHeader().setBackground(t.gutterBg);
        output.setBackground(t.dark ? t.bg.darker() : t.bg.darker());
        output.setBackground(blend(t.bg, t.dark ? Color.BLACK : Color.BLACK, 0.12f));
        output.setForeground(t.fg);
        output.setCaretColor(t.fg);
        moduleList.setBackground(t.gutterBg);
        moduleList.setForeground(t.fg);
        moduleList.setSelectionBackground(t.selection);
        moduleList.setSelectionForeground(t.dark ? Color.WHITE : Color.BLACK);
        immediate.setBackground(t.bg);
        immediate.setForeground(t.fg);
        immediate.setCaretColor(t.fg);
        immediate.setFont(new Font(editor.getFontFamily(), Font.PLAIN, 13));
        left.setBackground(t.gutterBg);
        bottom.setBackground(t.gutterBg);
        outHead.setBackground(t.gutterBg);
        immRow.setBackground(t.gutterBg);
        immLabel.setForeground(t.gutterCurrent);
        for (Component c : new Component[]{left.getComponent(0)}) c.setForeground(t.gutterCurrent);
        for (Component c : outHead.getComponents()) if (c instanceof JLabel) c.setForeground(t.gutterCurrent);
        left.getComponent(2).setBackground(t.gutterBg);
        output.setFont(new Font(editor.getFontFamily(), Font.PLAIN, Math.max(11, editor.getFontSize() - 1)));
        AppSettings.put("macro.theme", t.name);
        moduleList.repaint();
        repaint();
    }

    private static Color blend(Color a, Color b, float amount) {
        float inv = 1f - amount;
        return new Color(Math.round(a.getRed() * inv + b.getRed() * amount), Math.round(a.getGreen() * inv + b.getGreen() * amount),
                Math.round(a.getBlue() * inv + b.getBlue() * amount));
    }

    private void themeMenu(JButton anchor) {
        JPopupMenu m = new JPopupMenu();
        javax.swing.ButtonGroup g = new javax.swing.ButtonGroup();
        for (EditorTheme t : EditorTheme.ALL) {
            JRadioButtonMenuItem it = new JRadioButtonMenuItem(t.name, t == editor.getEditorTheme());
            it.addActionListener(e -> applyTheme(t));
            g.add(it);
            m.add(it);
        }
        m.show(anchor, 0, anchor.getHeight());
    }

    private void updateFontLabel() {
        fontLabel.setText(editor.getFontFamily() + "  " + editor.getFontSize() + " pt");
        output.setFont(new Font(editor.getFontFamily(), Font.PLAIN, Math.max(11, editor.getFontSize() - 1)));
        immediate.setFont(new Font(editor.getFontFamily(), Font.PLAIN, 13));
    }

    private void fontDialog() {
        List<String> fams = CodeFonts.monospaced();
        JList<String> list = new JList<>(fams.toArray(new String[0]));
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        list.setSelectedValue(editor.getFontFamily(), true);
        JSpinner size = new JSpinner(new SpinnerNumberModel(editor.getFontSize(), 8, 48, 1));
        JTextArea preview = new JTextArea("Sub Hello()\n    Dim n As Long          ' counter\n    For n = 1 To 10\n"
                + "        Cells(n, 1).Value = \"Row \" & n\n    Next n\nEnd Sub\n\n0O 1lI |!  {}[]()  <> =>  ->  ::");
        preview.setEditable(false);
        EditorTheme t = editor.getEditorTheme();
        preview.setBackground(t.bg);
        preview.setForeground(t.fg);
        preview.setMargin(new Insets(8, 10, 8, 10));
        Runnable upd = () -> {
            String f = list.getSelectedValue();
            if (f != null) preview.setFont(new Font(f, Font.PLAIN, (Integer) size.getValue()));
        };
        list.addListSelectionListener(e -> upd.run());
        size.addChangeListener(e -> upd.run());
        upd.run();
        JPanel p = new JPanel(new BorderLayout(10, 10));
        JScrollPane ls = new JScrollPane(list);
        ls.setPreferredSize(new Dimension(210, 260));
        JPanel lp = new JPanel(new BorderLayout(0, 6));
        lp.add(new JLabel("Font (monospaced fonts only)"), BorderLayout.NORTH);
        lp.add(ls, BorderLayout.CENTER);
        JPanel sp = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        sp.add(new JLabel("Size"));
        sp.add(size);
        lp.add(sp, BorderLayout.SOUTH);
        JScrollPane ps = new JScrollPane(preview);
        ps.setPreferredSize(new Dimension(440, 260));
        p.add(lp, BorderLayout.WEST);
        p.add(ps, BorderLayout.CENTER);
        int res = JOptionPane.showConfirmDialog(this, p, "Editor Font", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (res == JOptionPane.OK_OPTION && list.getSelectedValue() != null) {
            editor.setEditorFont(list.getSelectedValue(), (Integer) size.getValue());
            AppSettings.put("macro.font", editor.getFontFamily());
            AppSettings.put("macro.size", String.valueOf(editor.getFontSize()));
            updateFontLabel();
        }
    }

    // ------------------------------------------------------------------
    // Modules
    // ------------------------------------------------------------------

    /** (Re)reads the module list from the workbook - call after opening a file. */
    public void reloadModules() {
        flushQuiet();
        current = null;
        loadModules();
    }

    private void loadModules() {
        List<MacroModule> mods = host.modules();
        if (mods.isEmpty()) mods.add(new MacroModule("Module1", STARTER));
        moduleModel.clear();
        for (MacroModule m : mods) moduleModel.addElement(m);
        switchTo(mods.get(0));
        moduleList.setSelectedValue(mods.get(0), true);
        setTitle("Macro Studio — " + host.workbookName());
    }

    private void switchTo(MacroModule m) {
        flush();
        switching = true;
        try {
            current = m;
            editor.setCode(m.code);
            editor.goToLineStart();
        } finally {
            switching = false;
        }
        if (moduleList.getSelectedValue() != m) moduleList.setSelectedValue(m, true);
        refreshProcedures();
        updatePos();
        if (findBar.isVisible()) updateMatches();
    }

    /** Copies the editor text into the current module (and tells the app when it changed). */
    public void flush() {
        if (current == null) return;
        String code = editor.getCode();
        if (!code.equals(current.code)) {
            current.code = code.length() > MacroModule.MAX_CODE ? code.substring(0, MacroModule.MAX_CODE) : code;
            host.changed();
            moduleList.repaint();
        }
    }

    private void flushQuiet() {
        if (current != null) current.code = editor.getCode();
    }

    private String uniqueName(String base) {
        String n = base;
        int i = 2;
        while (nameTaken(n, null)) n = base + (i++);
        return n;
    }

    private boolean nameTaken(String n, MacroModule except) {
        for (MacroModule m : host.modules()) if (m != except && m.name.equalsIgnoreCase(n)) return true;
        return false;
    }

    private MacroModule addModule(String name, String code) {
        if (host.modules().size() >= MacroModule.MAX_MODULES) {
            JOptionPane.showMessageDialog(this, "A workbook can hold up to " + MacroModule.MAX_MODULES + " modules.", "Macro Studio", JOptionPane.WARNING_MESSAGE);
            return null;
        }
        flush();
        MacroModule m = new MacroModule(name, code);
        host.modules().add(m);
        moduleModel.addElement(m);
        host.changed();
        switchTo(m);
        return m;
    }

    private void newModule() {
        addModule(uniqueName("Module" + (host.modules().size() + 1)), "Sub NewMacro()\n    \nEnd Sub\n");
        editor.requestFocusInWindow();
    }

    private void renameModule() {
        if (current == null) return;
        String n = JOptionPane.showInputDialog(this, "Module name (letters, digits and _):", current.name);
        if (n == null) return;
        n = n.trim();
        if (!MacroModule.validName(n)) {
            JOptionPane.showMessageDialog(this, "Use a name that starts with a letter and has only letters, digits and _ (31 characters at most).", "Rename", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (nameTaken(n, current)) {
            JOptionPane.showMessageDialog(this, "There is already a module called " + n + ".", "Rename", JOptionPane.WARNING_MESSAGE);
            return;
        }
        current.name = n;
        host.changed();
        moduleList.repaint();
        refreshProcedures();
    }

    private void deleteModule() {
        if (current == null) return;
        if (host.modules().size() <= 1) {
            JOptionPane.showMessageDialog(this, "A workbook keeps at least one module. Clear its text instead.", "Delete module", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        int r = JOptionPane.showConfirmDialog(this, "Delete module " + current.name + "? This cannot be undone.", "Delete module",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);
        if (r != JOptionPane.OK_OPTION) return;
        MacroModule gone = current;
        int idx = host.modules().indexOf(gone);
        current = null;
        host.modules().remove(gone);
        moduleModel.removeElement(gone);
        host.changed();
        switchTo(host.modules().get(Math.max(0, Math.min(idx, host.modules().size() - 1))));
    }

    private void importBas() {
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Macro modules (*.bas, *.txt)", "bas", "txt", "vba"));
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        try {
            if (f.length() > MacroModule.MAX_CODE) {
                JOptionPane.showMessageDialog(this, "That file is too large for a macro module.", "Import", JOptionPane.WARNING_MESSAGE);
                return;
            }
            String text = new String(Files.readAllBytes(f.toPath()), StandardCharsets.UTF_8).replace("\r\n", "\n").replace('\r', '\n');
            String name = f.getName().replaceAll("\\.[^.]*$", "").replaceAll("[^A-Za-z0-9_]", "_");
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("(?m)^Attribute VB_Name = \"([^\"]+)\"\\n?").matcher(text);
            if (m.find()) {
                name = m.group(1);
                text = text.substring(0, m.start()) + text.substring(m.end());
            }
            if (!MacroModule.validName(name)) name = "Imported";
            addModule(uniqueName(name), text);
        } catch (java.io.IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not read the file:\n" + ex.getMessage(), "Import", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void exportBas() {
        if (current == null) return;
        flush();
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File(current.name + ".bas"));
        fc.setFileFilter(new FileNameExtensionFilter("Macro module (*.bas)", "bas"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        if (!f.getName().contains(".")) f = new File(f.getPath() + ".bas");
        if (f.exists() && JOptionPane.showConfirmDialog(this, f.getName() + " already exists. Replace it?", "Export",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE) != JOptionPane.OK_OPTION) return;
        try {
            String text = "Attribute VB_Name = \"" + current.name + "\"\r\n" + current.code.replace("\n", "\r\n");
            Files.write(f.toPath(), text.getBytes(StandardCharsets.UTF_8));
            infoLabel.setText("Saved " + f.getName());
        } catch (java.io.IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not save the file:\n" + ex.getMessage(), "Export", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void samplesMenu(JButton anchor) {
        JPopupMenu m = new JPopupMenu();
        for (String[] s : MacroSamples.ALL) {
            JMenuItem it = new JMenuItem(s[0]);
            it.addActionListener(e -> addModule(uniqueName(s[1]), s[2]));
            m.add(it);
        }
        m.show(anchor, 0, anchor.getHeight());
    }

    // ------------------------------------------------------------------
    // Procedures
    // ------------------------------------------------------------------

    private List<ProcRef> allProcedures() {
        List<ProcRef> out = new ArrayList<>();
        List<MacroModule> mods = host.modules();
        for (int i = 0; i < mods.size(); i++) {
            MacroModule m = mods.get(i);
            String code = m == current ? editor.getCode() : m.code;
            for (MacroEngine.ProcInfo p : MacroEngine.procedures(code)) out.add(new ProcRef(i, m.name, p));
        }
        return out;
    }

    private void refreshProcedures() {
        updatingCombo = true;
        try {
            Object sel = procCombo.getSelectedItem();
            procCombo.removeAllItems();
            int keep = -1;
            List<ProcRef> all = allProcedures();
            for (int i = 0; i < all.size(); i++) {
                procCombo.addItem(all.get(i));
                if (sel instanceof ProcRef && ((ProcRef) sel).info.name.equals(all.get(i).info.name)) keep = i;
            }
            if (keep >= 0) procCombo.setSelectedIndex(keep);
        } finally {
            updatingCombo = false;
        }
        infoLabel.setText(procCombo.getItemCount() + " procedure" + (procCombo.getItemCount() == 1 ? "" : "s") + " in the workbook");
        moduleList.repaint();
    }

    private List<VbaSyntax.Completion> otherModuleSymbols() {
        List<VbaSyntax.Completion> out = new ArrayList<>();
        for (MacroModule m : host.modules()) {
            if (m == current) continue;
            for (MacroEngine.ProcInfo p : MacroEngine.procedures(m.code)) out.add(new VbaSyntax.Completion(p.name, "P", p.function ? "function (" + m.name + ")" : "sub (" + m.name + ")"));
        }
        return out;
    }

    private void jumpTo(int moduleIdx, int line) {
        List<MacroModule> mods = host.modules();
        if (moduleIdx < 0 || moduleIdx >= mods.size()) return;
        if (mods.get(moduleIdx) != current) switchTo(mods.get(moduleIdx));
        editor.goToLine(line);
    }

    /** The procedure the cursor is inside (or null). */
    private String procedureAtCaret() {
        String[] lines = editor.getCode().split("\n", -1);
        int caret = editor.getCaretLine();
        String found = null;
        for (MacroEngine.ProcInfo p : MacroEngine.procedures(editor.getCode())) {
            if (p.line > caret) break;
            // find its end
            int end = lines.length;
            for (int i = p.line; i < lines.length; i++) {
                String t = lines[i].trim().toLowerCase(Locale.ROOT);
                if (t.startsWith("end sub") || t.startsWith("end function")) {
                    end = i + 1;
                    break;
                }
            }
            if (caret <= end) found = p.name;
        }
        return found;
    }

    // ------------------------------------------------------------------
    // Running
    // ------------------------------------------------------------------

    /** Names of Subs that can be run directly (no parameters), for the Macros dialog. */
    public List<String> runnableNames() {
        flushQuiet();
        List<String> out = new ArrayList<>();
        List<MacroModule> mods = host.modules();
        for (MacroModule m : mods) {
            for (MacroEngine.ProcInfo p : MacroEngine.procedures(m.code)) {
                if (!p.function && p.params == 0 && !p.isPrivate) out.add(p.name);
            }
        }
        Collections.sort(out, String.CASE_INSENSITIVE_ORDER);
        return out;
    }

    public void runCurrent() {
        flush();
        String name = procedureAtCaret();
        if (name == null) {
            ProcRef p = (ProcRef) procCombo.getSelectedItem();
            if (p != null) name = p.info.name;
        }
        if (name == null) {
            List<ProcRef> all = allProcedures();
            for (ProcRef p : all) {
                if (!p.info.function && p.info.params == 0) {
                    name = p.info.name;
                    break;
                }
            }
        }
        if (name == null) {
            print("There is no Sub to run. Write one like:  Sub Main() ... End Sub", 3);
            return;
        }
        runNamed(name);
    }

    /** Runs a procedure by name. @return null when it finished, otherwise the error text */
    public String runNamed(String name) {
        flush();
        SpreadsheetModel model = host.activeModel();
        if (model == null) {
            print("No active worksheet.", 3);
            return "No active worksheet.";
        }
        editor.setErrorLine(-1);
        print("▶ Running " + name + " ...", 0);
        long t0 = System.nanoTime();
        MacroEngine eng = newEngine(model);
        String error = null;
        try {
            List<String> codes = new ArrayList<>();
            for (MacroModule m : host.modules()) codes.add(m.code);
            eng.load(codes);
            int steps = eng.runProcedure(name);
            double ms = (System.nanoTime() - t0) / 1e6;
            print("✔ Finished " + name + " (" + steps + " steps, " + String.format(Locale.ROOT, "%.0f", ms) + " ms)", 2);
            infoLabel.setText("Last run: " + name + " ✓ " + String.format(Locale.ROOT, "%.0f", ms) + " ms");
        } catch (MacroEngine.MacroException ex) {
            error = ex.getMessage();
            reportError(ex);
        } catch (RuntimeException ex) {
            error = "Internal error: " + ex;
            print("✖ " + error, 1);
        }
        host.changed();
        return error;
    }

    private MacroEngine newEngine(SpreadsheetModel model) {
        MacroEngine eng = new MacroEngine(model, s -> print(s, 4));
        eng.setWorkbook(host.workbook());
        eng.setMsgBoxHandler(text -> {
            print("MsgBox: " + text, 5);
            JOptionPane.showMessageDialog(this, text, "Macro", JOptionPane.INFORMATION_MESSAGE);
        });
        eng.setInputBoxHandler(prompt -> JOptionPane.showInputDialog(this, prompt, "Macro", JOptionPane.QUESTION_MESSAGE));
        return eng;
    }

    private void reportError(MacroEngine.MacroException ex) {
        List<MacroModule> mods = host.modules();
        String where = ex.module >= 0 && ex.module < mods.size() && ex.line > 0
                ? "   [" + mods.get(ex.module).name + ", line " + ex.line + " — click to jump]" : "";
        printLink("✖ " + ex.getMessage() + where, ex.module, ex.line);
        if (ex.line > 0 && ex.module >= 0 && ex.module < mods.size()) {
            jumpTo(ex.module, ex.line);
            editor.setErrorLine(ex.line);
        }
    }

    private void runImmediate() {
        String text = immediate.getText().trim();
        if (text.isEmpty()) return;
        immediateHistory.add(text);
        historyPos = immediateHistory.size();
        immediate.setText("");
        print("› " + text, 5);
        SpreadsheetModel model = host.activeModel();
        if (model == null) return;
        flush();
        MacroEngine eng = newEngine(model);
        try {
            List<String> codes = new ArrayList<>();
            for (MacroModule m : host.modules()) codes.add(m.code);
            eng.load(codes);
            eng.immediate(text);
        } catch (MacroEngine.MacroException ex) {
            String msg = ex.getMessage().replaceFirst("^Line \\d+: ", "");
            print("✖ " + msg, 1);
        } catch (RuntimeException ex) {
            print("✖ Internal error: " + ex, 1);
        }
        host.changed();
    }

    /** kinds: 0 info, 1 error, 2 success, 3 warning, 4 program output, 5 message / echo. */
    private void print(String text, int kind) {
        append(text, kind, null);
    }

    private void printLink(String text, int module, int line) {
        append(text, 1, new int[]{module, line});
    }

    private void append(String text, int kind, int[] go) {
        EditorTheme t = editor.getEditorTheme();
        SimpleAttributeSet a = new SimpleAttributeSet();
        Color c;
        switch (kind) {
            case 1: c = t.dark ? new Color(0xFF6B6B) : new Color(0xC62828); break;
            case 2: c = t.dark ? new Color(0x6FD08C) : new Color(0x1B7F3B); break;
            case 3: c = t.dark ? new Color(0xFFB74D) : new Color(0xB26A00); break;
            case 5: c = t.kind[VbaSyntax.KEYWORD]; break;
            case 0: c = t.gutterFg.darker().equals(t.gutterFg) ? t.gutterFg : t.gutterCurrent; break;
            default: c = t.fg;
        }
        StyleConstants.setForeground(a, c);
        if (go != null) {
            StyleConstants.setUnderline(a, true);
            a.addAttribute("goto", go);
        }
        StyledDocument doc = output.getStyledDocument();
        try {
            doc.insertString(doc.getLength(), text + "\n", a);
            if (doc.getLength() > 200_000) doc.remove(0, doc.getLength() - 150_000);
        } catch (BadLocationException ignored) {
            // nothing
        }
        output.setCaretPosition(doc.getLength());
    }

    // ------------------------------------------------------------------
    // Find / go to
    // ------------------------------------------------------------------

    private void showFind(boolean replace) {
        findBar.setVisible(true);
        replRow.setVisible(replace);
        String sel = editor.getSelectedText();
        if (sel != null && !sel.isEmpty() && !sel.contains("\n")) findField.setText(sel);
        findField.requestFocusInWindow();
        findField.selectAll();
        updateMatches();
        revalidate();
    }

    private void hideFind() {
        findBar.setVisible(false);
        editor.clearMatches();
        editor.requestFocusInWindow();
        revalidate();
    }

    private void updateMatches() {
        int n = editor.markMatches(findField.getText(), caseBox.isSelected(), wordBox.isSelected(), regexBox.isSelected());
        findCount.setText(findField.getText().isEmpty() ? " " : n < 0 ? "bad pattern" : n == 0 ? "no matches" : n + (n == 1 ? " match" : " matches"));
    }

    private void gotoLineDialog() {
        String s = JOptionPane.showInputDialog(this, "Go to line (1 - " + editor.getLineCount() + "):", editor.getCaretLine());
        if (s == null) return;
        try {
            editor.goToLine(Integer.parseInt(s.trim()));
        } catch (NumberFormatException ex) {
            // ignore
        }
    }

    private void updatePos() {
        posLabel.setText("Ln " + editor.getCaretLine() + ", Col " + editor.getCaretColumn()
                + (current == null ? "" : "   │   " + current.name));
    }

    // ------------------------------------------------------------------
    // Help
    // ------------------------------------------------------------------

    private void help() {
        String html = "<html><body style='font-family:sans-serif;font-size:12px;margin:6px'>"
                + "<h2>Macro Studio</h2>"
                + "<p>Write macros in a practical subset of VBA. Press <b>F5</b> to run the Sub under the cursor. "
                + "Macros are saved inside the workbook (.spx) and never run by themselves when a file is opened.</p>"
                + "<h3>Language</h3><ul>"
                + "<li><b>Procedures:</b> Sub / Function (Public, Private), parameters (ByVal, ByRef, Optional), Call, Exit Sub / Exit Function</li>"
                + "<li><b>Variables:</b> Dim x As Long, Dim a(1 To 10), Dim g(5,5), ReDim [Preserve], Const, Set</li>"
                + "<li><b>Decisions:</b> If / ElseIf / Else / End If, one-line If, Select Case (values, lists, To, Is &gt; 5, Case Else)</li>"
                + "<li><b>Loops:</b> For / Next [Step], For Each, Do While / Until ... Loop, Do ... Loop While / Until, While ... Wend, Exit For / Exit Do</li>"
                + "<li><b>Other:</b> With ... End With, On Error Resume Next, Err.Number / Err.Description, End</li>"
                + "<li><b>Operators:</b> + - * / \\ Mod ^ &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;= And Or Not Xor Like</li></ul>"
                + "<h3>Worksheet objects</h3><ul>"
                + "<li>Range(\"A1:B3\"), Cells(r, c), ActiveCell, Selection, ActiveSheet, Worksheets(\"Name\"), Rows, Columns, UsedRange</li>"
                + "<li>.Value .Formula .Text .Address .Row .Column .Count .Rows.Count .Columns.Count .Offset(r,c) .Resize(r,c) .End(xlUp)</li>"
                + "<li>.Font.Bold / Italic / Size / Color / Name, .Interior.Color, .HorizontalAlignment, .NumberFormat, .ColumnWidth</li>"
                + "<li>.ClearContents, .Clear, Application.WorksheetFunction.Sum / Average / Count / Max / Min / Median / CountIf / SumIf</li></ul>"
                + "<h3>Functions</h3><p>Abs Int Fix Sqr Rnd Round Sgn Exp Log Sin Cos Tan | Len Left Right Mid Trim UCase LCase InStr Replace Space String Chr Asc StrReverse "
                + "| CStr CInt CLng CDbl CBool Val IsNumeric IsEmpty | Format Now Date Time Timer Year Month Day DateAdd DateDiff | Split Join Array UBound LBound | "
                + "RGB MsgBox InputBox Evaluate</p>"
                + "<h3>Shortcuts</h3><table cellpadding='2'>"
                + "<tr><td><b>F5</b></td><td>Run</td><td><b>Ctrl+Space</b></td><td>Completion list</td></tr>"
                + "<tr><td><b>Ctrl+Shift+F</b></td><td>Format code</td><td><b>Ctrl+/</b></td><td>Comment / uncomment</td></tr>"
                + "<tr><td><b>Tab / Shift+Tab</b></td><td>Indent / outdent</td><td><b>Ctrl+D</b></td><td>Duplicate line</td></tr>"
                + "<tr><td><b>Alt+Up / Down</b></td><td>Move line</td><td><b>Ctrl+Shift+K</b></td><td>Delete line</td></tr>"
                + "<tr><td><b>Ctrl+F / Ctrl+H</b></td><td>Find / replace</td><td><b>Ctrl+G</b></td><td>Go to line</td></tr>"
                + "<tr><td><b>Ctrl+wheel, Ctrl +/-</b></td><td>Zoom</td><td><b>Ctrl+0</b></td><td>Reset zoom</td></tr></table>"
                + "<p>The <b>Immediate</b> box under the output runs one statement; start with <b>?</b> to print a value, e.g. <tt>? Sum1(3)</tt>.</p>"
                + "<p><i>Limits:</i> macros stop after 20 seconds or 5 million steps; GoTo/labels, classes, UserForms and file access are not supported.</p>"
                + "</body></html>";
        JEditorPane pane = new JEditorPane("text/html", html);
        pane.setEditable(false);
        JScrollPane sp = new JScrollPane(pane);
        sp.setPreferredSize(new Dimension(680, 520));
        JOptionPane.showMessageDialog(this, sp, "Macro Studio Help", JOptionPane.PLAIN_MESSAGE);
        pane.setCaretPosition(0);
    }
}
