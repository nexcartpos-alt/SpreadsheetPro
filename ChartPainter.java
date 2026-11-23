import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.Shape;
import java.awt.geom.Arc2D;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;

/**
 * Draws a chart in its own coordinate space (0,0)-(w,h). The caller translates and scales the
 * graphics, so charts follow the sheet's zoom. While painting it records the clickable / hoverable
 * shapes ({@link Hit}) so the layer can show tooltips and react to legend clicks.
 */
public final class ChartPainter {

    public static final Color[] PALETTE = {
            new Color(0x2E7D32), new Color(0x1E88E5), new Color(0xFB8C00), new Color(0x8E24AA),
            new Color(0xE53935), new Color(0x00ACC1), new Color(0xFDD835), new Color(0x6D4C41)};
    private static Color TEXT = new Color(0x404040);
    private static Color GRID = new Color(0xE3E3E3);
    private static Color AXIS = new Color(0x9E9E9E);
    private static Color BG = Color.WHITE;
    private static Color FRAME = new Color(0xBDBDBD);

    /** Picks the light or dark chart colours (exports and printing always use the light ones). */
    private static void applyTheme(boolean dark) {
        if (dark) {
            TEXT = new Color(0xDADADA); GRID = new Color(0x3C4146); AXIS = new Color(0x8C9399);
            BG = new Color(0x22262A); FRAME = new Color(0x4A5056);
        } else {
            TEXT = new Color(0x404040); GRID = new Color(0xE3E3E3); AXIS = new Color(0x9E9E9E);
            BG = Color.WHITE; FRAME = new Color(0xBDBDBD);
        }
    }

    /** A shape the mouse can point at. */
    public static final class Hit {
        public final Shape shape;
        public final String tip;
        public final int series;
        public final int point;
        public final boolean legend;

        Hit(Shape shape, String tip, int series, int point, boolean legend) {
            this.shape = shape;
            this.tip = tip;
            this.series = series;
            this.point = point;
            this.legend = legend;
        }
    }

    private ChartPainter() {
    }

    public static Color color(int i) {
        return PALETTE[Math.floorMod(i, PALETTE.length)];
    }

    // ------------------------------------------------------------------

    private static final class Axis {
        double lo, hi, step;

        double pos(double v, double from, double len) {
            return from + (v - lo) / (hi - lo) * len;
        }
    }

    private static Axis niceAxis(double min, double max, boolean includeZero) {
        if (includeZero) {
            min = Math.min(min, 0);
            max = Math.max(max, 0);
        }
        if (!(max > min)) {
            max = min + 1;
        }
        double raw = (max - min) / 5.0;
        double mag = Math.pow(10, Math.floor(Math.log10(raw)));
        double norm = raw / mag;
        double step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
        Axis a = new Axis();
        a.step = step;
        a.lo = Math.floor(min / step + 1e-9) * step;
        a.hi = Math.ceil(max / step - 1e-9) * step;
        if (a.hi <= a.lo) a.hi = a.lo + step;
        return a;
    }

    private static String tick(double v, double step) {
        if (Math.abs(v) < step * 1e-9) v = 0;
        int dec = step >= 1 ? 0 : (int) Math.min(6, Math.ceil(-Math.log10(step) - 1e-9));
        double a = Math.abs(v);
        if (a >= 1e9) return trim(v / 1e9, dec) + "B";
        if (a >= 1e6) return trim(v / 1e6, Math.max(dec, 0)) + "M";
        return String.format("%,." + dec + "f", v);
    }

    private static String trim(double v, int dec) {
        String s = String.format("%." + Math.max(1, dec) + "f", v);
        return s.contains(".") ? s.replaceAll("0+$", "").replaceAll("\\.$", "") : s;
    }

    /** Value as shown in tooltips. */
    public static String num(double v) {
        return FormulaEngine.formatNumber(v);
    }

    private static String fit(String s, FontMetrics fm, int maxWidth) {
        if (fm.stringWidth(s) <= maxWidth) return s;
        String t = s;
        while (t.length() > 1 && fm.stringWidth(t + "…") > maxWidth) t = t.substring(0, t.length() - 1);
        return t + "…";
    }

    // ------------------------------------------------------------------

    /**
     * @param anim         0..1 entrance animation
     * @param hoverSeries  highlighted series (-1 none)
     * @param hoverPoint   highlighted point (-1 none)
     * @param hits         receives the shapes in paint order (may be null)
     */
    public static void paint(Graphics2D g0, ChartSpec s, ChartData d, int w, int h, double anim,
                             int hoverSeries, int hoverPoint, List<Hit> hits) {
        paint(g0, s, d, w, h, anim, hoverSeries, hoverPoint, hits, false);
    }

    public static void paint(Graphics2D g0, ChartSpec s, ChartData d, int w, int h, double anim,
                             int hoverSeries, int hoverPoint, List<Hit> hits, boolean dark) {
        Graphics2D g = (Graphics2D) g0.create();
        applyTheme(dark);
        try {
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_PURE);
            g.setColor(BG);
            g.fillRect(0, 0, w, h);
            g.setColor(FRAME);
            g.drawRect(0, 0, w - 1, h - 1);

            int top = 10;
            if (s.title != null && !s.title.isEmpty()) {
                g.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 14));
                FontMetrics fm = g.getFontMetrics();
                String t = fit(s.title, fm, w - 20);
                g.setColor(TEXT);
                g.drawString(t, (w - fm.stringWidth(t)) / 2, top + fm.getAscent());
                top += fm.getHeight() + 6;
            }
            if (d.series() == 0 || d.points == 0 || d.isEmpty()) {
                g.setFont(new Font(Font.SANS_SERIF, Font.ITALIC, 12));
                g.setColor(AXIS);
                String msg = "No numbers in " + s.rangeText();
                g.drawString(msg, (w - g.getFontMetrics().stringWidth(msg)) / 2, h / 2);
                return;
            }
            boolean pie = s.type == ChartSpec.Type.PIE || s.type == ChartSpec.Type.DOUGHNUT;
            int legendCount = pie ? d.points : d.series();
            boolean showLegend = s.legend && (pie || d.series() > 1);
            int bottom = 8;
            List<Object[]> legendItems = new ArrayList<>(); // {label, colorIndex, Rectangle}
            if (showLegend) {
                bottom += layoutLegend(g, s, d, pie, legendCount, w, legendItems, h - 8);
            }
            Rectangle plot = new Rectangle(10, top, w - 22, Math.max(20, h - top - bottom));
            switch (s.type) {
                case PIE:
                    paintPie(g, s, d, plot, anim, hoverPoint, hits, false);
                    break;
                case DOUGHNUT:
                    paintPie(g, s, d, plot, anim, hoverPoint, hits, true);
                    break;
                case BAR:
                    paintBars(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, true, false, false);
                    break;
                case STACKED_BAR:
                    paintBars(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, true, true, false);
                    break;
                case STACKED_COLUMN:
                    paintBars(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, false, true, false);
                    break;
                case COLUMN:
                    paintBars(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, false, false, false);
                    break;
                case SCATTER:
                    paintLines(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, false, true);
                    break;
                case RADAR:
                    paintRadar(g, s, d, plot, anim, hoverSeries, hoverPoint, hits);
                    break;
                case COMBO:
                    paintBars(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, false, false, true);
                    break;
                default:
                    paintLines(g, s, d, plot, anim, hoverSeries, hoverPoint, hits, s.type == ChartSpec.Type.AREA, false);
            }
            if (showLegend) paintLegend(g, s, pie, legendItems, hits);
        } finally {
            g.dispose();
        }
    }

    /** The finished chart as a picture ({@code scale} = pixels per chart pixel). */
    public static java.awt.image.BufferedImage toImage(ChartSpec s, ChartData d, double scale) {
        int w = (int) Math.round(s.width * scale);
        int h = (int) Math.round(s.height * scale);
        java.awt.image.BufferedImage img = new java.awt.image.BufferedImage(w, h, java.awt.image.BufferedImage.TYPE_INT_RGB);
        Graphics2D g = img.createGraphics();
        try {
            g.scale(scale, scale);
            paint(g, s, d, s.width, s.height, 1.0, -1, -1, null);
        } finally {
            g.dispose();
        }
        return img;
    }

    // ---- legend ----

    private static int layoutLegend(Graphics2D g, ChartSpec s, ChartData d, boolean pie, int count, int w,
                                    List<Object[]> items, int bottomY) {
        g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        FontMetrics fm = g.getFontMetrics();
        int rowH = 18;
        List<Object[]> row = new ArrayList<>();
        List<List<Object[]>> rows = new ArrayList<>();
        int rowW = 0;
        for (int i = 0; i < count; i++) {
            String label = fit(pie ? d.categories.get(i) : d.names.get(i), fm, Math.max(40, w / 3));
            int iw = 14 + fm.stringWidth(label) + 14;
            if (!row.isEmpty() && rowW + iw > w - 20 && rows.size() < 2) {
                rows.add(row);
                row = new ArrayList<>();
                rowW = 0;
            }
            if (rows.size() >= 2 && rowW + iw > w - 20) break; // too many entries: the rest are not listed
            row.add(new Object[]{label, i, iw});
            rowW += iw;
        }
        if (!row.isEmpty()) rows.add(row);
        int y = bottomY - rows.size() * rowH;
        for (List<Object[]> r : rows) {
            int total = 0;
            for (Object[] it : r) total += (Integer) it[2];
            int x = (w - total) / 2;
            for (Object[] it : r) {
                items.add(new Object[]{it[0], it[1], new Rectangle(x, y, (Integer) it[2], rowH)});
                x += (Integer) it[2];
            }
            y += rowH;
        }
        return rows.size() * rowH + 4;
    }

    private static void paintLegend(Graphics2D g, ChartSpec s, boolean pie, List<Object[]> items, List<Hit> hits) {
        g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        FontMetrics fm = g.getFontMetrics();
        for (Object[] it : items) {
            int idx = (Integer) it[1];
            Rectangle r = (Rectangle) it[2];
            boolean off = s.hidden.contains(idx);
            g.setColor(off ? new Color(0xC8C8C8) : color(idx));
            g.fillRoundRect(r.x + 2, r.y + 5, 8, 8, 3, 3);
            g.setColor(off ? new Color(0xA0A0A0) : TEXT);
            g.drawString((String) it[0], r.x + 14, r.y + (r.height + fm.getAscent() - fm.getDescent()) / 2);
            if (off) {
                int sw = fm.stringWidth((String) it[0]);
                int ly = r.y + r.height / 2;
                g.drawLine(r.x + 14, ly, r.x + 14 + sw, ly);
            }
            if (hits != null) {
                hits.add(new Hit(r, off ? "Click to show" : "Click to hide", pie ? -1 : idx, pie ? idx : -1, true));
            }
        }
    }

    // ---- bar / column ----

    private static void paintBars(Graphics2D g, ChartSpec s, ChartData d, Rectangle area, double anim,
                                  int hs, int hp, List<Hit> hits, boolean horizontal, boolean stacked, boolean combo) {
        List<Integer> all = visibleSeries(s, d);
        List<Integer> vis = all;
        List<Integer> lineVis = new ArrayList<>();
        if (combo && all.size() > 1) { // the last series becomes a line over the columns
            vis = new ArrayList<>(all.subList(0, all.size() - 1));
            lineVis.add(all.get(all.size() - 1));
        }
        double min = 0;
        double max = 0;
        if (stacked) {
            for (int i = 0; i < d.points; i++) {
                double pos = 0;
                double neg = 0;
                for (int si : vis) {
                    double v = d.values.get(si)[i];
                    if (Double.isNaN(v)) continue;
                    if (v >= 0) pos += v;
                    else neg += v;
                }
                max = Math.max(max, pos);
                min = Math.min(min, neg);
            }
        } else {
            for (int si : all) for (double v : d.values.get(si)) if (!Double.isNaN(v)) {
                min = Math.min(min, v);
                max = Math.max(max, v);
            }
        }
        Axis ax = niceAxis(min, max, true);
        g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        FontMetrics fm = g.getFontMetrics();

        int catLabelW = 0;
        if (horizontal) {
            for (String c : d.categories) catLabelW = Math.max(catLabelW, fm.stringWidth(c));
            catLabelW = Math.min(catLabelW, area.width / 4) + 8;
        }
        int valLabelW = 0;
        if (!horizontal) {
            for (double v = ax.lo; v <= ax.hi + ax.step / 2; v += ax.step) {
                valLabelW = Math.max(valLabelW, fm.stringWidth(tick(v, ax.step)));
            }
            valLabelW += 8;
        }
        Rectangle p = horizontal
                ? new Rectangle(area.x + catLabelW, area.y + 4, area.width - catLabelW - 6, area.height - fm.getHeight() - 10)
                : new Rectangle(area.x + valLabelW, area.y + 6, area.width - valLabelW - 4, area.height - fm.getHeight() - 12);

        // grid + value labels
        g.setStroke(new BasicStroke(1f));
        for (double v = ax.lo; v <= ax.hi + ax.step / 2; v += ax.step) {
            String t = tick(v, ax.step);
            g.setColor(GRID);
            if (horizontal) {
                int x = (int) Math.round(ax.pos(v, p.x, p.width));
                g.drawLine(x, p.y, x, p.y + p.height);
                g.setColor(TEXT);
                g.drawString(t, x - fm.stringWidth(t) / 2, p.y + p.height + fm.getAscent() + 3);
            } else {
                int y = (int) Math.round(p.y + p.height - (v - ax.lo) / (ax.hi - ax.lo) * p.height);
                g.drawLine(p.x, y, p.x + p.width, y);
                g.setColor(TEXT);
                g.drawString(t, p.x - 6 - fm.stringWidth(t), y + (fm.getAscent() - fm.getDescent()) / 2);
            }
        }
        double zero = horizontal ? ax.pos(0, p.x, p.width) : p.y + p.height - (0 - ax.lo) / (ax.hi - ax.lo) * p.height;
        g.setColor(AXIS);
        if (horizontal) g.drawLine((int) Math.round(zero), p.y, (int) Math.round(zero), p.y + p.height);
        else g.drawLine(p.x, (int) Math.round(zero), p.x + p.width, (int) Math.round(zero));

        int n = d.points;
        int ns = Math.max(1, vis.size());
        double band = (horizontal ? p.height : p.width) / (double) n;
        double gap = band * 0.18;
        double barW = stacked ? band - gap * 2 : (band - gap * 2) / ns;
        int labelEvery = Math.max(1, (int) Math.ceil(fm.getHeight() * 1.3 / Math.max(1, band)));
        for (int i = 0; i < n; i++) {
            double start = (horizontal ? p.y : p.x) + band * i;
            // category label
            g.setColor(TEXT);
            if (horizontal) {
                if (i % labelEvery == 0) {
                    String t = fit(d.categories.get(i), fm, catLabelW - 8);
                    g.drawString(t, p.x - 6 - fm.stringWidth(t), (int) (start + band / 2 + (fm.getAscent() - fm.getDescent()) / 2.0));
                }
            } else {
                if (i % labelEvery == 0) {
                    int maxW = (int) (band * labelEvery) - 4;
                    String t = fit(d.categories.get(i), fm, Math.max(10, maxW));
                    g.drawString(t, (int) (start + band / 2.0 - fm.stringWidth(t) / 2.0), p.y + p.height + fm.getAscent() + 4);
                }
            }
            double posBase = 0;
            double negBase = 0;
            for (int k = 0; k < vis.size(); k++) {
                int si = vis.get(k);
                double v = d.values.get(si)[i];
                if (Double.isNaN(v)) continue;
                double off = stacked ? start + gap : start + gap + barW * k;
                double v0 = 0;
                double v1 = v;
                if (stacked) {
                    v0 = v >= 0 ? posBase : negBase;
                    v1 = v0 + v;
                    if (v >= 0) posBase = v1;
                    else negBase = v1;
                }
                double a0 = zeroToValue(v0 * anim, ax, p, horizontal);
                double a1 = zeroToValue(v1 * anim, ax, p, horizontal);
                Rectangle2D r;
                if (horizontal) {
                    r = new Rectangle2D.Double(Math.min(a0, a1), off, Math.max(1, Math.abs(a1 - a0)), Math.max(1, barW - 1));
                } else {
                    r = new Rectangle2D.Double(off, Math.min(a0, a1), Math.max(1, barW - 1), Math.max(1, Math.abs(a1 - a0)));
                }
                boolean hot = si == hs && i == hp;
                g.setColor(hot ? color(si).brighter() : color(si));
                g.fill(r);
                if (stacked) {
                    g.setColor(BG);
                    g.setStroke(new BasicStroke(1f));
                    g.draw(r);
                }
                if (hot) {
                    g.setColor(color(si).darker());
                    g.setStroke(new BasicStroke(2f));
                    g.draw(r);
                    g.setStroke(new BasicStroke(1f));
                }
                if (s.dataLabels && anim > 0.95) {
                    String t = num(v);
                    int tw = fm.stringWidth(t);
                    g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 10));
                    FontMetrics sm = g.getFontMetrics();
                    tw = sm.stringWidth(t);
                    if (stacked) {
                        if (r.getWidth() >= tw + 4 && r.getHeight() >= sm.getHeight()) {
                            g.setColor(Color.WHITE);
                            g.drawString(t, (float) (r.getCenterX() - tw / 2.0),
                                    (float) (r.getCenterY() + (sm.getAscent() - sm.getDescent()) / 2.0));
                        }
                    } else if (horizontal) {
                        g.setColor(TEXT);
                        g.drawString(t, (float) (v >= 0 ? r.getMaxX() + 3 : r.getMinX() - tw - 3),
                                (float) (r.getCenterY() + (sm.getAscent() - sm.getDescent()) / 2.0));
                    } else {
                        g.setColor(TEXT);
                        g.drawString(t, (float) (r.getCenterX() - tw / 2.0),
                                (float) (v >= 0 ? r.getMinY() - 3 : r.getMaxY() + sm.getAscent()));
                    }
                    g.setFont(fm.getFont());
                }
                if (hits != null) {
                    hits.add(new Hit(r, d.categories.get(i) + "\n" + d.names.get(si) + ": " + num(v), si, i, false));
                }
            }
        }
        for (int si : lineVis) {
            Color c = color(si);
            double[] vals = d.values.get(si);
            java.awt.geom.Path2D.Double path = new java.awt.geom.Path2D.Double();
            boolean started = false;
            for (int i = 0; i < n; i++) {
                if (Double.isNaN(vals[i])) { started = false; continue; }
                double x = p.x + band * (i + 0.5);
                double y = p.y + p.height - (vals[i] * anim - ax.lo) / (ax.hi - ax.lo) * p.height;
                if (!started) { path.moveTo(x, y); started = true; } else path.lineTo(x, y);
            }
            g.setColor(c);
            g.setStroke(new BasicStroke(2.5f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g.draw(path);
            g.setStroke(new BasicStroke(1f));
            for (int i = 0; i < n; i++) {
                if (Double.isNaN(vals[i])) continue;
                double x = p.x + band * (i + 0.5);
                double y = p.y + p.height - (vals[i] * anim - ax.lo) / (ax.hi - ax.lo) * p.height;
                boolean hot = si == hs && i == hp;
                double r = hot ? 6.5 : 4;
                g.setColor(BG);
                g.fill(new Ellipse2D.Double(x - r - 1, y - r - 1, r * 2 + 2, r * 2 + 2));
                g.setColor(hot ? c.darker() : c);
                Ellipse2D dot = new Ellipse2D.Double(x - r, y - r, r * 2, r * 2);
                g.fill(dot);
                if (hits != null) {
                    hits.add(new Hit(new Ellipse2D.Double(x - 8, y - 8, 16, 16),
                            d.categories.get(i) + "\n" + d.names.get(si) + ": " + num(vals[i]), si, i, false));
                }
            }
        }
        if (!horizontal && !stacked && !combo) {
            double[] px = new double[n];
            double[] xv = new double[n];
            for (int i = 0; i < n; i++) {
                px[i] = p.x + band * (i + 0.5);
                xv[i] = i + 1;
            }
            paintTrends(g, s, d, vis, xv, px, p, ax.lo, ax.hi, anim);
        }
    }

    // ---- trendlines ----

    /** Draws the trendline of every visible series; px/xv are pixel and data x of each point. */
    private static void paintTrends(Graphics2D g, ChartSpec s, ChartData d, List<Integer> vis, double[] xv, double[] px,
                                    Rectangle p, double lo, double hi, double anim) {
        if (s.trend == ChartSpec.Trend.NONE || anim < 0.95) return;
        java.awt.Shape old = g.getClip();
        g.clipRect(p.x, p.y - 2, p.width, p.height + 4);
        List<String[]> notes = new ArrayList<>();
        List<Color> noteColors = new ArrayList<>();
        for (int si : vis) {
            double[] ys = d.values.get(si);
            Color c = color(si).darker();
            g.setColor(c);
            g.setStroke(new BasicStroke(1.8f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_ROUND, 10f, new float[]{7f, 5f}, 0f));
            if (s.trend == ChartSpec.Trend.MOVING_AVERAGE) {
                double[] ma = Trendline.movingAverage(ys, d.points >= 4 ? 3 : 2);
                java.awt.geom.Path2D.Double path = new java.awt.geom.Path2D.Double();
                boolean started = false;
                for (int i = 0; i < ma.length; i++) {
                    if (Double.isNaN(ma[i]) || Double.isNaN(px[i])) continue;
                    double y = p.y + p.height - (ma[i] - lo) / (hi - lo) * p.height;
                    if (!started) {
                        path.moveTo(px[i], y);
                        started = true;
                    } else {
                        path.lineTo(px[i], y);
                    }
                }
                if (started) g.draw(path);
                continue;
            }
            Trendline t = Trendline.fit(s.trend, xv, ys);
            if (t == null) continue;
            // pixel position is linear in x: find two points with different x to get the mapping
            int i0 = -1;
            int i1 = -1;
            for (int i = 0; i < xv.length; i++) {
                if (Double.isNaN(xv[i]) || Double.isNaN(px[i])) continue;
                if (i0 < 0 || xv[i] < xv[i0]) i0 = i;
                if (i1 < 0 || xv[i] > xv[i1]) i1 = i;
            }
            if (i0 < 0 || xv[i1] == xv[i0]) continue;
            double scale = (px[i1] - px[i0]) / (xv[i1] - xv[i0]);
            java.awt.geom.Path2D.Double path = new java.awt.geom.Path2D.Double();
            boolean started = false;
            for (int k = 0; k <= 48; k++) {
                double x = xv[i0] + (xv[i1] - xv[i0]) * k / 48.0;
                double y = p.y + p.height - (t.at(x) - lo) / (hi - lo) * p.height;
                double X = px[i0] + (x - xv[i0]) * scale;
                if (!started) {
                    path.moveTo(X, y);
                    started = true;
                } else {
                    path.lineTo(X, y);
                }
            }
            g.draw(path);
            notes.add(new String[]{(d.series() > 1 ? d.names.get(si) + ":  " : "") + t.label()});
            noteColors.add(c);
        }
        g.setClip(old);
        if (s.showEquation && !notes.isEmpty()) {
            g.setFont(new Font(Font.SANS_SERIF, Font.ITALIC, 10));
            FontMetrics fm = g.getFontMetrics();
            int y = p.y + fm.getAscent() + 2;
            for (int k = 0; k < notes.size(); k++) {
                String t = fit(notes.get(k)[0], fm, p.width - 12);
                int w = fm.stringWidth(t);
                g.setColor(new Color(BG.getRed(), BG.getGreen(), BG.getBlue(), 210));
                g.fillRect(p.x + 6, y - fm.getAscent(), w + 6, fm.getHeight());
                g.setColor(noteColors.get(k));
                g.drawString(t, p.x + 9, y);
                y += fm.getHeight();
            }
        }
        g.setStroke(new BasicStroke(1f));
    }

    private static double zeroToValue(double v, Axis ax, Rectangle p, boolean horizontal) {
        return horizontal ? ax.pos(v, p.x, p.width) : p.y + p.height - (v - ax.lo) / (ax.hi - ax.lo) * p.height;
    }

    // ---- line / area / scatter ----

    private static void paintLines(Graphics2D g, ChartSpec s, ChartData d, Rectangle area, double anim,
                                   int hs, int hp, List<Hit> hits, boolean fillArea, boolean scatter) {
        List<Integer> vis = visibleSeries(s, d);
        double min = Double.MAX_VALUE;
        double max = -Double.MAX_VALUE;
        for (int si : vis) for (double v : d.values.get(si)) if (!Double.isNaN(v)) {
            min = Math.min(min, v);
            max = Math.max(max, v);
        }
        if (min == Double.MAX_VALUE) {
            min = 0;
            max = 1;
        }
        Axis ay = niceAxis(min, max, fillArea || (min >= 0 && min < max * 0.6));
        Axis axx = null;
        boolean numericX = scatter && d.xs != null;
        if (numericX) {
            double xmin = Double.MAX_VALUE;
            double xmax = -Double.MAX_VALUE;
            for (double x : d.xs) if (!Double.isNaN(x)) {
                xmin = Math.min(xmin, x);
                xmax = Math.max(xmax, x);
            }
            if (xmin == Double.MAX_VALUE) {
                xmin = 0;
                xmax = 1;
            }
            axx = niceAxis(xmin, xmax, false);
        }
        g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        FontMetrics fm = g.getFontMetrics();
        int valLabelW = 0;
        for (double v = ay.lo; v <= ay.hi + ay.step / 2; v += ay.step) {
            valLabelW = Math.max(valLabelW, fm.stringWidth(tick(v, ay.step)));
        }
        valLabelW += 8;
        Rectangle p = new Rectangle(area.x + valLabelW, area.y + 6, area.width - valLabelW - 8, area.height - fm.getHeight() - 12);

        g.setStroke(new BasicStroke(1f));
        for (double v = ay.lo; v <= ay.hi + ay.step / 2; v += ay.step) {
            int y = (int) Math.round(p.y + p.height - (v - ay.lo) / (ay.hi - ay.lo) * p.height);
            g.setColor(GRID);
            g.drawLine(p.x, y, p.x + p.width, y);
            g.setColor(TEXT);
            String t = tick(v, ay.step);
            g.drawString(t, p.x - 6 - fm.stringWidth(t), y + (fm.getAscent() - fm.getDescent()) / 2);
        }
        g.setColor(AXIS);
        g.drawLine(p.x, p.y + p.height, p.x + p.width, p.y + p.height);

        int n = d.points;
        double[] px = new double[n];
        double band = p.width / (double) Math.max(1, n);
        for (int i = 0; i < n; i++) {
            if (numericX) {
                px[i] = Double.isNaN(d.xs[i]) ? Double.NaN : axx.pos(d.xs[i], p.x, p.width);
            } else {
                px[i] = p.x + band * (i + 0.5);
            }
        }
        // x labels
        g.setColor(TEXT);
        if (numericX) {
            for (double v = axx.lo; v <= axx.hi + axx.step / 2; v += axx.step) {
                String t = tick(v, axx.step);
                int x = (int) Math.round(axx.pos(v, p.x, p.width));
                g.setColor(GRID);
                g.drawLine(x, p.y, x, p.y + p.height);
                g.setColor(TEXT);
                g.drawString(t, x - fm.stringWidth(t) / 2, p.y + p.height + fm.getAscent() + 4);
            }
        } else {
            int every = Math.max(1, (int) Math.ceil(60.0 / Math.max(1, band)));
            for (int i = 0; i < n; i += every) {
                String t = fit(d.categories.get(i), fm, (int) Math.max(14, band * every - 4));
                g.drawString(t, (int) (px[i] - fm.stringWidth(t) / 2.0), p.y + p.height + fm.getAscent() + 4);
            }
        }

        Shape oldClip = g.getClip();
        if (!scatter) g.clipRect(p.x - 2, p.y - 8, (int) Math.ceil((p.width + 14) * anim), p.height + 20);
        for (int k = 0; k < vis.size(); k++) {
            int si = vis.get(k);
            double[] vals = d.values.get(si);
            Color c = color(si);
            java.awt.geom.Path2D.Double path = new java.awt.geom.Path2D.Double();
            boolean started = false;
            double baseY = p.y + p.height - (Math.max(ay.lo, Math.min(0, ay.hi)) - ay.lo) / (ay.hi - ay.lo) * p.height;
            double firstX = 0;
            double lastX = 0;
            for (int i = 0; i < n; i++) {
                if (Double.isNaN(vals[i]) || Double.isNaN(px[i])) continue;
                double y = p.y + p.height - (vals[i] - ay.lo) / (ay.hi - ay.lo) * p.height;
                if (!started) {
                    path.moveTo(px[i], y);
                    started = true;
                    firstX = px[i];
                } else {
                    path.lineTo(px[i], y);
                }
                lastX = px[i];
            }
            if (started && !scatter) {
                if (fillArea) {
                    java.awt.geom.Path2D.Double fillPath = new java.awt.geom.Path2D.Double(path);
                    fillPath.lineTo(lastX, baseY);
                    fillPath.lineTo(firstX, baseY);
                    fillPath.closePath();
                    g.setColor(new Color(c.getRed(), c.getGreen(), c.getBlue(), 70));
                    g.fill(fillPath);
                }
                g.setColor(c);
                g.setStroke(new BasicStroke(si == hs ? 3.2f : 2.2f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g.draw(path);
            }
            g.setStroke(new BasicStroke(1f));
            for (int i = 0; i < n; i++) {
                if (Double.isNaN(vals[i]) || Double.isNaN(px[i])) continue;
                double y = p.y + p.height - (vals[i] - ay.lo) / (ay.hi - ay.lo) * p.height;
                boolean hot = si == hs && i == hp;
                double r = (hot ? 6.5 : 4) * (scatter ? Math.max(0.2, anim) : 1);
                Ellipse2D dot = new Ellipse2D.Double(px[i] - r, y - r, r * 2, r * 2);
                g.setColor(BG);
                g.fill(new Ellipse2D.Double(px[i] - r - 1, y - r - 1, r * 2 + 2, r * 2 + 2));
                g.setColor(hot ? c.darker() : c);
                g.fill(dot);
                if (hits != null) {
                    String head = numericX ? "x = " + num(d.xs[i]) : d.categories.get(i);
                    hits.add(new Hit(new Ellipse2D.Double(px[i] - 9, y - 9, 18, 18),
                            head + "\n" + d.names.get(si) + ": " + num(vals[i]), si, i, false));
                }
            }
        }
        g.setClip(oldClip);
        {
            double[] xv = new double[n];
            for (int i = 0; i < n; i++) xv[i] = numericX ? d.xs[i] : i + 1;
            paintTrends(g, s, d, vis, xv, px, p, ay.lo, ay.hi, anim);
        }
        if (s.dataLabels && anim > 0.95) {
            g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 10));
            FontMetrics lf = g.getFontMetrics();
            g.setColor(TEXT);
            int every = Math.max(1, (int) Math.ceil(34.0 / Math.max(1, band)));
            for (int si : vis) {
                double[] vals = d.values.get(si);
                for (int i = 0; i < n; i += every) {
                    if (Double.isNaN(vals[i]) || Double.isNaN(px[i])) continue;
                    double y = p.y + p.height - (vals[i] - ay.lo) / (ay.hi - ay.lo) * p.height;
                    String t = num(vals[i]);
                    g.drawString(t, (float) (px[i] - lf.stringWidth(t) / 2.0), (float) (y - 8));
                }
            }
        }
    }

    // ---- pie ----

    private static void paintPie(Graphics2D g, ChartSpec s, ChartData d, Rectangle area, double anim,
                                 int hp, List<Hit> hits, boolean doughnut) {
        double[] vals = d.values.get(0);
        double total = 0;
        for (int i = 0; i < vals.length; i++) {
            if (!s.hidden.contains(i) && !Double.isNaN(vals[i]) && vals[i] > 0) total += vals[i];
        }
        if (total <= 0) {
            g.setColor(AXIS);
            g.setFont(new Font(Font.SANS_SERIF, Font.ITALIC, 12));
            g.drawString("Nothing to show", area.x + area.width / 2 - 40, area.y + area.height / 2);
            return;
        }
        int size = Math.min(area.width, area.height) - 24;
        if (size < 20) size = 20;
        double cx = area.x + area.width / 2.0;
        double cy = area.y + area.height / 2.0;
        double angle = 90;
        g.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 11));
        FontMetrics fm = g.getFontMetrics();
        for (int i = 0; i < vals.length; i++) {
            if (s.hidden.contains(i) || Double.isNaN(vals[i]) || vals[i] <= 0) continue;
            double sweep = -360.0 * vals[i] / total;
            double mid = Math.toRadians(angle + sweep / 2);
            boolean hot = i == hp;
            double ox = hot ? Math.cos(mid) * 8 : 0;
            double oy = hot ? -Math.sin(mid) * 8 : 0;
            Arc2D.Double arc = new Arc2D.Double(cx - size / 2.0 + ox, cy - size / 2.0 + oy, size, size,
                    angle, sweep * anim, Arc2D.PIE);
            java.awt.geom.Area shape = new java.awt.geom.Area(arc);
            if (doughnut) {
                double hs = size * 0.52;
                shape.subtract(new java.awt.geom.Area(new Ellipse2D.Double(cx - hs / 2 + ox, cy - hs / 2 + oy, hs, hs)));
            }
            g.setColor(color(i));
            g.fill(shape);
            g.setColor(BG);
            g.setStroke(new BasicStroke(1.5f));
            g.draw(shape);
            double frac = vals[i] / total;
            if ((frac > 0.06 || (s.dataLabels && frac > 0.03)) && anim > 0.95) {
                String t = s.dataLabels ? num(vals[i]) : Math.round(frac * 100) + "%";
                g.setColor(Color.WHITE);
                g.drawString(t, (int) (cx + ox + Math.cos(mid) * size * (doughnut ? 0.38 : 0.32) - fm.stringWidth(t) / 2.0),
                        (int) (cy + oy - Math.sin(mid) * size * (doughnut ? 0.38 : 0.32) + fm.getAscent() / 2.0 - 1));
            }
            if (hits != null) {
                hits.add(new Hit(shape, d.categories.get(i) + "\n" + d.names.get(0) + ": " + num(vals[i])
                        + "  (" + String.format("%.1f", frac * 100) + "%)", 0, i, false));
            }
            angle += sweep;
        }
    }

    // ---- radar ----

    private static void paintRadar(Graphics2D g, ChartSpec s, ChartData d, Rectangle area, double anim,
                                   int hs, int hp, List<Hit> hits) {
        List<Integer> vis = visibleSeries(s, d);
        double max = 0;
        for (int si : vis) for (double v : d.values.get(si)) if (!Double.isNaN(v)) max = Math.max(max, v);
        Axis ax = niceAxis(0, max <= 0 ? 1 : max, true);
        g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        FontMetrics fm = g.getFontMetrics();
        int n = Math.max(3, d.points);
        int labelPad = fm.getHeight() + 10;
        double cx = area.x + area.width / 2.0;
        double cy = area.y + area.height / 2.0;
        double rad = Math.max(10, Math.min(area.width / 2.0 - 40, area.height / 2.0 - labelPad));
        g.setStroke(new BasicStroke(1f));
        for (double v = ax.lo + ax.step; v <= ax.hi + ax.step / 2; v += ax.step) { // rings
            double rr = (v - ax.lo) / (ax.hi - ax.lo) * rad;
            java.awt.geom.Path2D.Double ring = new java.awt.geom.Path2D.Double();
            for (int i = 0; i < n; i++) {
                double a = -Math.PI / 2 + 2 * Math.PI * i / n;
                double x = cx + Math.cos(a) * rr, y = cy + Math.sin(a) * rr;
                if (i == 0) ring.moveTo(x, y); else ring.lineTo(x, y);
            }
            ring.closePath();
            g.setColor(GRID);
            g.draw(ring);
            String t = tick(v, ax.step);
            g.setColor(AXIS);
            g.drawString(t, (int) (cx + 3), (int) (cy - rr + fm.getAscent() - 1));
        }
        for (int i = 0; i < n; i++) { // spokes + category names
            double a = -Math.PI / 2 + 2 * Math.PI * i / n;
            g.setColor(GRID);
            g.draw(new java.awt.geom.Line2D.Double(cx, cy, cx + Math.cos(a) * rad, cy + Math.sin(a) * rad));
            if (i < d.points) {
                String t = fit(d.categories.get(i), fm, 90);
                double lx = cx + Math.cos(a) * (rad + 8), ly = cy + Math.sin(a) * (rad + 8);
                double tw = fm.stringWidth(t);
                double x = Math.cos(a) > 0.3 ? lx : Math.cos(a) < -0.3 ? lx - tw : lx - tw / 2;
                double y = Math.sin(a) > 0.3 ? ly + fm.getAscent() : Math.sin(a) < -0.3 ? ly : ly + fm.getAscent() / 2.0;
                g.setColor(TEXT);
                g.drawString(t, (float) x, (float) y);
            }
        }
        for (int si : vis) {
            double[] vals = d.values.get(si);
            Color c = color(si);
            java.awt.geom.Path2D.Double poly = new java.awt.geom.Path2D.Double();
            double[] pxs = new double[d.points], pys = new double[d.points];
            for (int i = 0; i < d.points; i++) {
                double v = Double.isNaN(vals[i]) ? 0 : vals[i];
                double rr = (v * anim - ax.lo) / (ax.hi - ax.lo) * rad;
                double a = -Math.PI / 2 + 2 * Math.PI * i / n;
                pxs[i] = cx + Math.cos(a) * rr;
                pys[i] = cy + Math.sin(a) * rr;
                if (i == 0) poly.moveTo(pxs[i], pys[i]); else poly.lineTo(pxs[i], pys[i]);
            }
            poly.closePath();
            g.setColor(new Color(c.getRed(), c.getGreen(), c.getBlue(), 55));
            g.fill(poly);
            g.setColor(c);
            g.setStroke(new BasicStroke(2f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g.draw(poly);
            g.setStroke(new BasicStroke(1f));
            for (int i = 0; i < d.points; i++) {
                boolean hot = si == hs && i == hp;
                double r = hot ? 6 : 3.5;
                g.setColor(BG);
                g.fill(new Ellipse2D.Double(pxs[i] - r - 1, pys[i] - r - 1, r * 2 + 2, r * 2 + 2));
                g.setColor(hot ? c.darker() : c);
                g.fill(new Ellipse2D.Double(pxs[i] - r, pys[i] - r, r * 2, r * 2));
                if (hits != null && !Double.isNaN(vals[i])) {
                    hits.add(new Hit(new Ellipse2D.Double(pxs[i] - 8, pys[i] - 8, 16, 16),
                            d.categories.get(i) + "\n" + d.names.get(si) + ": " + num(vals[i]), si, i, false));
                }
            }
        }
    }

    private static List<Integer> visibleSeries(ChartSpec s, ChartData d) {
        List<Integer> vis = new ArrayList<>();
        for (int i = 0; i < d.series(); i++) if (!s.hidden.contains(i)) vis.add(i);
        return vis;
    }
}
