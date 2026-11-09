import java.io.File;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

/** Prints the same JSON shape as cmd/xlsxdump for Apache POI. */
public class PoiDump {
    public static void main(String[] args) throws Exception {
        if (args.length != 1) {
            System.err.println("usage: PoiDump file.xlsx");
            System.exit(2);
        }
        StringBuilder b = new StringBuilder();
        try (Workbook wb = WorkbookFactory.create(new File(args[0]))) {
            boolean date1904 = wb instanceof org.apache.poi.ss.usermodel.Date1904Support
                    && ((org.apache.poi.ss.usermodel.Date1904Support) wb).isDate1904();
            b.append("{\"date1904\":").append(date1904);
            b.append(",\"names\":[");
            List<String> names = new ArrayList<>();
            for (Name n : wb.getAllNames()) {
                String ref = "";
                try {
                    ref = n.getRefersToFormula();
                } catch (Exception ignored) {
                }
                if (ref == null) ref = "";
                names.add("{\"name\":" + jq(n.getNameName()) + ",\"ref\":" + jq(ref) + "}");
            }
            Collections.sort(names);
            for (int i = 0; i < names.size(); i++) {
                if (i > 0) b.append(',');
                b.append(names.get(i));
            }
            b.append("],\"sheets\":[");
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                if (s > 0) b.append(',');
                Sheet sh = wb.getSheetAt(s);
                b.append("{\"name\":").append(jq(sh.getSheetName()));
                String state = "visible";
                if (wb.isSheetVeryHidden(s)) state = "veryHidden";
                else if (wb.isSheetHidden(s)) state = "hidden";
                b.append(",\"state\":").append(jq(state));
                b.append(",\"merges\":[");
                List<String> merges = new ArrayList<>();
                for (CellRangeAddress a : sh.getMergedRegions()) {
                    merges.add(a.formatAsString());
                }
                Collections.sort(merges);
                for (int i = 0; i < merges.size(); i++) {
                    if (i > 0) b.append(',');
                    b.append(jq(merges.get(i)));
                }
                b.append("],\"cells\":[");
                List<String> cells = new ArrayList<>();
                List<String> comments = new ArrayList<>();
                List<String> links = new ArrayList<>();
                for (Row row : sh) {
                    for (Cell cell : row) {
                        String addr = cell.getAddress().formatAsString();
                        String formula = "";
                        CellType t = cell.getCellType();
                        if (t == CellType.FORMULA) {
                            formula = dataTableFormula(cell);
                            if (formula.isEmpty()) {
                                try {
                                    formula = cell.getCellFormula();
                                } catch (Exception ex) {
                                    formula = "";
                                }
                            }
                        }
                        String typ = "s";
                        String val = null;
                        boolean hasVal = false;
                        CellType vt = t == CellType.FORMULA ? cell.getCachedFormulaResultType() : t;
                        try {
                            if (vt == CellType.NUMERIC) {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    typ = "d";
                                    LocalDateTime ldt = cell.getLocalDateTimeCellValue();
                                    if (ldt != null) {
                                        val = jq(ldt.atOffset(ZoneOffset.UTC).format(DateTimeFormatter.ISO_OFFSET_DATE_TIME));
                                        hasVal = true;
                                    }
                                } else {
                                    typ = "n";
                                    val = Double.toString(cell.getNumericCellValue());
                                    hasVal = true;
                                }
                            } else if (vt == CellType.BOOLEAN) {
                                typ = "b";
                                val = cell.getBooleanCellValue() ? "true" : "false";
                                hasVal = true;
                            } else if (vt == CellType.STRING) {
                                typ = "s";
                                val = jq(cell.getStringCellValue());
                                hasVal = true;
                            } else if (vt == CellType.ERROR) {
                                typ = "e";
                                val = jq(org.apache.poi.ss.usermodel.FormulaError.forInt(cell.getErrorCellValue()).getString());
                                hasVal = true;
                            }
                        } catch (Exception ignored) {
                        }
                        if (formula.isEmpty() && !hasVal) {
                            continue;
                        }
                        StringBuilder c = new StringBuilder();
                        c.append("{\"a\":").append(jq(addr));
                        c.append(",\"t\":").append(jq(formula.isEmpty() ? typ : "f"));
                        if (hasVal) {
                            c.append(",\"v\":").append(val);
                        }
                        if (!formula.isEmpty()) {
                            c.append(",\"f\":").append(jq(formula));
                        }
                        String fmt = "";
                        try {
                            if (cell.getCellStyle() != null) {
                                fmt = cell.getCellStyle().getDataFormatString();
                            }
                        } catch (Exception ignored) {
                        }
                        if (fmt != null && !fmt.isEmpty() && !"General".equals(fmt)) {
                            c.append(",\"nf\":").append(jq(fmt));
                        }
                        c.append('}');
                        cells.add(c.toString());
                        Comment note = cell.getCellComment();
                        if (note != null && note.getString() != null) {
                            String author = note.getAuthor() == null ? "" : note.getAuthor();
                            comments.add("{\"a\":" + jq(addr) + ",\"author\":" + jq(author) + ",\"text\":" + jq(note.getString().getString()) + "}");
                        }
                        Hyperlink link = cell.getHyperlink();
                        if (link != null) {
                            String url = link.getAddress();
                            if (url == null || url.isEmpty()) url = link.getLabel();
                            if (url != null && !url.isEmpty()) {
                                links.add("{\"a\":" + jq(addr) + ",\"url\":" + jq(url) + "}");
                            }
                        }
                    }
                }
                for (int i = 0; i < cells.size(); i++) {
                    if (i > 0) b.append(',');
                    b.append(cells.get(i));
                }
                b.append("],\"comments\":[");
                for (int i = 0; i < comments.size(); i++) {
                    if (i > 0) b.append(',');
                    b.append(comments.get(i));
                }
                b.append("],\"links\":[");
                for (int i = 0; i < links.size(); i++) {
                    if (i > 0) b.append(',');
                    b.append(links.get(i));
                }
                b.append("]}");
            }
            b.append("]}");
        } catch (Exception ex) {
            System.out.println("{\"error\":" + jq(ex.toString()) + ",\"date1904\":false,\"names\":[],\"sheets\":[]}");
            System.exit(1);
        }
        System.out.println(b);
    }

    static String dataTableFormula(Cell cell) {
        if (!(cell instanceof org.apache.poi.xssf.usermodel.XSSFCell)) {
            return "";
        }
        org.apache.poi.xssf.usermodel.XSSFCell xc = (org.apache.poi.xssf.usermodel.XSSFCell) cell;
        if (xc.getCTCell() == null || !xc.getCTCell().isSetF()) {
            return "";
        }
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula f = xc.getCTCell().getF();
        if (f.getT() == null || f.getT().intValue() != org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType.INT_DATA_TABLE) {
            return "";
        }
        String r1 = f.getR1() == null ? "" : f.getR1();
        String r2 = f.getR2() == null ? "" : f.getR2();
        if (!r2.isEmpty()) {
            return "TABLE(" + r1 + "," + r2 + ")";
        }
        if (!r1.isEmpty()) {
            return "TABLE(" + r1 + ")";
        }
        return "TABLE()";
    }

    static String jq(String s) {
        if (s == null) s = "";
        StringBuilder b = new StringBuilder();
        b.append('"');
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '\\': b.append("\\\\"); break;
                case '"': b.append("\\\""); break;
                case '\n': b.append("\\n"); break;
                case '\r': b.append("\\r"); break;
                case '\t': b.append("\\t"); break;
                default:
                    if (c < 0x20) b.append(String.format("\\u%04x", (int) c));
                    else b.append(c);
            }
        }
        b.append('"');
        return b.toString();
    }
}
