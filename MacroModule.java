/** One module of macro code stored in the workbook (like a VBA standard module). */
public final class MacroModule {
    public static final int MAX_CODE = 2_000_000;
    public static final int MAX_MODULES = 100;

    public String name;
    public String code;

    public MacroModule(String name, String code) {
        this.name = name;
        this.code = code == null ? "" : code;
    }

    /** Letters, digits and underscore, starting with a letter. */
    public static boolean validName(String n) {
        return n != null && n.matches("[A-Za-z][A-Za-z0-9_]{0,30}");
    }

    @Override
    public String toString() {
        return name;
    }
}
