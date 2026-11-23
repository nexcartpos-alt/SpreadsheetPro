import java.awt.Color;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;
import javax.swing.SwingConstants;

/** Ready-made workbooks for File > New > Templates. */
public final class Templates {
    public static final class Template {
        public final String name;
        public final String description;
        public final String icon;
        private final Supplier<List<WorkbookIO.SheetData>> maker;

        Template(String name, String description, String icon, Supplier<List<WorkbookIO.SheetData>> maker) {
            this.name = name;
            this.description = description;
            this.icon = icon;
            this.maker = maker;
        }

        public List<WorkbookIO.SheetData> create() {
            return maker.get();
        }
    }

    private static final Color GREEN = new Color(0x1E7145);
    private static final Color LIGHT = new Color(0xE2F0D9);
    private static final Color INPUT = new Color(0xFFF2CC);

    public static final List<Template> ALL = new ArrayList<>();

    static {
        ALL.add(new Template("Monthly budget", "Income, expenses and what is left each month", "table", Templates::budget));
        ALL.add(new Template("Invoice", "Items, quantity, price, tax and total", "save", Templates::invoice));
        ALL.add(new Template("Loan calculator", "Monthly payment and total interest (PMT)", "calc", Templates::loan));
        ALL.add(new Template("Gradebook", "Scores, average and a letter grade", "table", Templates::gradebook));
        ALL.add(new Template("Attendance sheet", "Mark P / A for each day and count them", "table", Templates::attendance));
        ALL.add(new Template("Project tracker", "Tasks, owner, status and due date with a drop-down", "table", Templates::tracker));
    }

    private Templates() {
    }

    /** Small helper to fill a sheet. */
    private static final class Sheet {
        final SpreadsheetModel m = new SpreadsheetModel();

        Sheet set(int r, int c, String text) {
            m.setInput(r, c, text);
            return this;
        }

        Sheet title(int r, int c, String text) {
            m.setInput(r, c, text);
            m.updateStyle(r, c, s -> {
                s.bold = true;
                s.fontSize = 18;
                s.foreground = GREEN;
            });
            return this;
        }

        Sheet header(int r, int c1, int c2, String... names) {
            for (int c = c1; c <= c2; c++) {
                final int cc = c;
                m.setInput(r, c, names[c - c1]);
                m.updateStyle(r, c, s -> {
                    s.bold = true;
                    s.background = GREEN;
                    s.foreground = Color.WHITE;
                    s.alignment = SwingConstants.CENTER;
                });
            }
            return this;
        }

        Sheet bold(int r, int c) {
            m.updateStyle(r, c, s -> s.bold = true);
            return this;
        }

        Sheet money(int r1, int c1, int r2, int c2) {
            for (int r = r1; r <= r2; r++) for (int c = c1; c <= c2; c++) m.updateStyle(r, c, s -> {
                s.numFmt = CellFormats.CURRENCY;
                s.decimals = 2;
            });
            return this;
        }

        Sheet fill(int r1, int c1, int r2, int c2, Color bg) {
            for (int r = r1; r <= r2; r++) for (int c = c1; c <= c2; c++) m.updateStyle(r, c, s -> s.background = bg);
            return this;
        }

        Sheet width(int c, int w) {
            m.setColumnWidth(c, w);
            return this;
        }
    }

    private static List<WorkbookIO.SheetData> one(String name, Sheet s) {
        List<WorkbookIO.SheetData> l = new ArrayList<>();
        l.add(new WorkbookIO.SheetData(name, s.m));
        return l;
    }

    private static List<WorkbookIO.SheetData> budget() {
        Sheet s = new Sheet();
        s.title(0, 0, "Monthly budget");
        s.header(2, 0, 3, "Item", "Planned", "Actual", "Difference");
        s.set(3, 0, "INCOME").bold(3, 0);
        String[] inc = {"Salary", "Other income"};
        int r = 4;
        for (String n : inc) {
            s.set(r, 0, n).set(r, 1, "0").set(r, 2, "0").set(r, 3, "=C" + (r + 1) + "-B" + (r + 1));
            r++;
        }
        s.set(r, 0, "Total income").bold(r, 0).set(r, 1, "=SUM(B5:B6)").set(r, 2, "=SUM(C5:C6)").set(r, 3, "=C7-B7");
        s.fill(r, 0, r, 3, LIGHT);
        int incTotal = r + 1; // row number 7
        s.set(8, 0, "EXPENSES").bold(8, 0);
        String[] exp = {"Rent", "Groceries", "Transport", "Utilities", "Phone / Internet", "Insurance", "Fun", "Savings"};
        r = 9;
        for (String n : exp) {
            s.set(r, 0, n).set(r, 1, "0").set(r, 2, "0").set(r, 3, "=B" + (r + 1) + "-C" + (r + 1));
            r++;
        }
        s.set(r, 0, "Total expenses").bold(r, 0).set(r, 1, "=SUM(B10:B" + r + ")").set(r, 2, "=SUM(C10:C" + r + ")")
                .set(r, 3, "=B" + (r + 1) + "-C" + (r + 1));
        s.fill(r, 0, r, 3, LIGHT);
        int expTotal = r + 1;
        r += 2;
        s.set(r, 0, "LEFT OVER").bold(r, 0).set(r, 1, "=B" + incTotal + "-B" + expTotal).set(r, 2, "=C" + incTotal + "-C" + expTotal);
        s.fill(r, 0, r, 3, LIGHT);
        s.money(4, 1, r, 3);
        s.fill(4, 1, 5, 2, INPUT).fill(9, 1, 16, 2, INPUT);
        s.m.addCondRule(CondRule.dataBar(9, 2, 16, 2, "63BE7B"));
        s.width(0, 160).width(1, 100).width(2, 100).width(3, 100);
        return one("Budget", s);
    }

    private static List<WorkbookIO.SheetData> invoice() {
        Sheet s = new Sheet();
        s.title(0, 0, "INVOICE");
        s.set(2, 0, "Bill to:").bold(2, 0).set(2, 1, "Customer name");
        s.set(3, 0, "Invoice #").bold(3, 0).set(3, 1, "1001");
        s.set(4, 0, "Date").bold(4, 0).set(4, 1, "=TODAY()");
        s.header(6, 0, 4, "Description", "Qty", "Unit price", "Tax %", "Amount");
        for (int r = 7; r < 14; r++) {
            int n = r + 1;
            s.set(r, 4, "=IF(B" + n + "=\"\",\"\",B" + n + "*C" + n + "*(1+D" + n + "))");
            s.fill(r, 0, r, 3, INPUT);
        }
        s.set(7, 0, "Example item").set(7, 1, "2").set(7, 2, "49.99").set(7, 3, "0.1");
        s.m.updateStyle(7, 3, st -> st.numFmt = CellFormats.PERCENT);
        for (int r = 8; r < 14; r++) s.m.updateStyle(r, 3, st -> st.numFmt = CellFormats.PERCENT);
        s.money(7, 2, 14, 2).money(7, 4, 14, 4);
        s.set(14, 3, "TOTAL").bold(14, 3).set(14, 4, "=SUM(E8:E14)").bold(14, 4).money(14, 4, 14, 4);
        s.fill(14, 3, 14, 4, LIGHT);
        s.width(0, 220).width(1, 70).width(2, 100).width(3, 70).width(4, 110);
        return one("Invoice", s);
    }

    private static List<WorkbookIO.SheetData> loan() {
        Sheet s = new Sheet();
        s.title(0, 0, "Loan calculator");
        s.set(2, 0, "Loan amount").set(2, 1, "20000").fill(2, 1, 2, 1, INPUT).money(2, 1, 2, 1);
        s.set(3, 0, "Yearly interest rate").set(3, 1, "0.065").fill(3, 1, 3, 1, INPUT);
        s.m.updateStyle(3, 1, st -> {
            st.numFmt = CellFormats.PERCENT;
            st.decimals = 2;
        });
        s.set(4, 0, "Years").set(4, 1, "5").fill(4, 1, 4, 1, INPUT);
        s.set(6, 0, "Monthly payment").bold(6, 0).set(6, 1, "=-PMT(B4/12,B5*12,B3)").money(6, 1, 6, 1).bold(6, 1);
        s.set(7, 0, "Total paid").set(7, 1, "=B7*B5*12").money(7, 1, 7, 1);
        s.set(8, 0, "Total interest").set(8, 1, "=B8-B3").money(8, 1, 8, 1);
        s.fill(6, 0, 8, 1, LIGHT);
        s.width(0, 170).width(1, 120);
        return one("Loan", s);
    }

    private static List<WorkbookIO.SheetData> gradebook() {
        Sheet s = new Sheet();
        s.title(0, 0, "Gradebook");
        s.header(2, 0, 5, "Student", "Test 1", "Test 2", "Test 3", "Average", "Grade");
        String[] names = {"Student 1", "Student 2", "Student 3", "Student 4", "Student 5", "Student 6", "Student 7", "Student 8"};
        for (int i = 0; i < names.length; i++) {
            int r = 3 + i;
            int n = r + 1;
            s.set(r, 0, names[i]).set(r, 1, String.valueOf(60 + (i * 7) % 35)).set(r, 2, String.valueOf(55 + (i * 11) % 40))
                    .set(r, 3, String.valueOf(70 + (i * 5) % 30));
            s.set(r, 4, "=ROUND(AVERAGE(B" + n + ":D" + n + "),1)");
            s.set(r, 5, "=IF(E" + n + ">=90,\"A\",IF(E" + n + ">=80,\"B\",IF(E" + n + ">=70,\"C\",IF(E" + n + ">=60,\"D\",\"F\"))))");
            s.fill(r, 1, r, 3, INPUT);
        }
        s.set(12, 0, "Class average").bold(12, 0);
        for (int c = 1; c <= 4; c++) {
            String col = FormulaEngine.columnLabel(c);
            s.set(12, c, "=ROUND(AVERAGE(" + col + "4:" + col + "11),1)").bold(12, c);
        }
        s.m.addCondRule(CondRule.colorScale(3, 4, 10, 4, "F8696B", "FFEB84", "63BE7B"));
        s.width(0, 130);
        return one("Grades", s);
    }

    private static List<WorkbookIO.SheetData> attendance() {
        Sheet s = new Sheet();
        s.title(0, 0, "Attendance");
        String[] hdr = new String[9];
        hdr[0] = "Name";
        for (int i = 1; i <= 5; i++) hdr[i] = "Day " + i;
        hdr[6] = "Present";
        hdr[7] = "Absent";
        hdr[8] = "%";
        s.header(2, 0, 8, hdr);
        for (int i = 0; i < 10; i++) {
            int r = 3 + i;
            int n = r + 1;
            s.set(r, 0, "Student " + (i + 1));
            for (int c = 1; c <= 5; c++) {
                s.set(r, c, (i + c) % 5 == 0 ? "A" : "P");
                s.m.updateStyle(r, c, st -> st.alignment = SwingConstants.CENTER);
            }
            s.set(r, 6, "=COUNTIF(B" + n + ":F" + n + ",\"P\")").set(r, 7, "=COUNTIF(B" + n + ":F" + n + ",\"A\")")
                    .set(r, 8, "=G" + n + "/5");
            s.m.updateStyle(r, 8, st -> st.numFmt = CellFormats.PERCENT);
        }
        s.m.addCondRule(new CondRule(CondRule.Type.EQUAL, 3, 1, 12, 5, "A", "", new Color(0xFFC7CE), new Color(0x9C0006)));
        s.width(0, 130);
        return one("Attendance", s);
    }

    private static List<WorkbookIO.SheetData> tracker() {
        Sheet s = new Sheet();
        s.title(0, 0, "Project tracker");
        s.header(2, 0, 4, "Task", "Owner", "Status", "Due date", "Done?");
        String[][] rows = {{"Plan the project", "Ali", "Done", "2026-12-01"}, {"Build first draft", "Sara", "In progress", "2026-12-10"},
            {"Review", "Omar", "Not started", "2026-12-15"}, {"Publish", "Ali", "Not started", "2026-12-20"}};
        for (int i = 0; i < rows.length; i++) {
            int r = 3 + i;
            for (int c = 0; c < 4; c++) s.set(r, c, rows[i][c]);
            s.set(r, 4, "=IF(C" + (r + 1) + "=\"Done\",\"Yes\",\"\")");
        }
        s.m.addValidation(new ValidationRule(ValidationRule.Type.LIST, null, 3, 2, 30, 2, "Not started,In progress,Done", "",
                true, true, ValidationRule.Alert.STOP, "", ""));
        s.m.addCondRule(new CondRule(CondRule.Type.EQUAL, 3, 2, 30, 2, "Done", "", new Color(0xC6EFCE), new Color(0x006100)));
        s.m.addCondRule(new CondRule(CondRule.Type.EQUAL, 3, 2, 30, 2, "In progress", "", new Color(0xFFEB9C), new Color(0x9C5700)));
        s.width(0, 180).width(1, 90).width(2, 110).width(3, 100);
        return one("Tracker", s);
    }
}
