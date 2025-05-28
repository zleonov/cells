package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;
import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

import java.lang.reflect.Field;
import java.util.Comparator;
import java.util.Optional;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;

import com.google.common.collect.Streams;

import software.leonov.cells.Workbooks.Format;

/**
 * Static methods for working with {@link Sheet}s.
 * 
 * @author Zhenya Leonov
 */
final public class Sheets {

    private Sheets() {
    }

    /**
     * Enables filtering for the range of cells covering the top row horizontally and the entire sheet vertically.
     * 
     * @param sheet the specified sheet
     * @return the specified sheet
     */
    public static Sheet autoFilterTopRow(final Sheet sheet) {
        getFirstRowIndex(sheet).ifPresent(firstRow -> {
            final Row row = sheet.getRow(firstRow);
            Rows.getFirstCellIndex(row).ifPresent(firstCell -> {
                Rows.getLastCellIndex(row).ifPresent(lastCell -> {
                    final int lastRow = Format.of(sheet.getWorkbook()).getMaxRowNum() - 1;
                    autoFilter(sheet, firstRow, firstCell, lastRow, lastCell);
                });
            });
        });

        return sheet;
    }

    /**
     * Returns the index (0-based) of the first row in the specified sheet or an empty {@code Optional} if sheet has no
     * defined rows.
     * 
     * @param sheet the specified sheet
     * @return the index (0-based) of the first row in the specified sheet or an empty {@code Optional} if sheet has no
     *         defined rows
     */
    public static Optional<Integer> getFirstRowIndex(final Sheet sheet) {
        final int firstRowIndex = sheet.getFirstRowNum();
        return firstRowIndex < 0 ? Optional.empty() : Optional.of(firstRowIndex);
    }

    /**
     * Enables filtering for the given range of cells in the specified sheet.
     * 
     * @param sheet   the specified sheet
     * @param fromRow the 0-based index of the the first row
     * @param fromCol the 0-based index of the first column
     * @param toRow   the 0-based index of the the last row
     * @param toCol   the 0-based index of the last column
     * @return the specified sheet
     */
    public static Sheet autoFilter(final Sheet sheet, final int fromRow, final int fromCol, final int toRow, final int toCol) {
        checkNotNull(sheet, "sheet == null");
        sheet.setAutoFilter(new CellRangeAddress(fromRow, toRow, fromCol, toCol));
        return sheet;
    }

    /**
     * Enables filtering for the given range of cells in the specified sheet.
     * 
     * @param sheet   the specified sheet
     * @param fromRow the 0-based index of the the first row
     * @param fromCol the letter reference of first column
     * @param toRow   the 0-based index of the the last row
     * @param toCol   the letter reference of the last column
     * @return the specified sheet
     */
    public static Sheet autoFilter(final Sheet sheet, final int fromRow, final String fromCol, final int toRow, final String toCol) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(fromCol, "fromCol == null");
        checkNotNull(toCol, "toCol == null");
        return autoFilter(sheet, fromRow, toRow, convertColStringToIndex(fromCol), convertColStringToIndex(toCol));
    }

//    /**
//     * Enables filtering for the given range of cells in the specified sheet.
//     * 
//     * @param sheet    the specified sheet
//     * @param firstRow the first row
//     * @param firstCol the 0-based index of the first column
//     * @param lastRow  the last row
//     * @param lastCol  the 0-based index of the last column
//     * @return the specified sheet
//     */
//    public static Sheet autoFilter(final Sheet sheet, final Row firstRow, final int firstCol, final Row lastRow, final int lastCol) {
//        checkNotNull(sheet, "sheet == null");
//        checkNotNull(firstRow, "firstRow == null");
//        checkNotNull(lastRow, "lastRow == null");
//        sheet.setAutoFilter(new CellRangeAddress(firstRow.getRowNum(), lastRow.getRowNum(), firstCol, lastCol));
//        return sheet;
//    }
//
//    /**
//     * Enables filtering for the given range of cells in the specified sheet.
//     * 
//     * @param sheet    the specified sheet
//     * @param firstRow the first row
//     * @param firstCol the letter reference of first column
//     * @param lastRow  the last row
//     * @param lastCol  the letter reference of the last column
//     * @return the specified sheet
//     */
//    public static Sheet autoFilter(final Sheet sheet, final Row firstRow, final String firstCol, final Row lastRow, final String lastCol) {
//        checkNotNull(sheet, "sheet == null");
//        checkNotNull(firstRow, "firstRow == null");
//        checkNotNull(firstCol, "firstCol == null");
//        checkNotNull(lastRow, "lastRow == null");
//        checkNotNull(lastCol, "lastCol == null");
//        sheet.setAutoFilter(new CellRangeAddress(firstRow.getRowNum(), lastRow.getRowNum(), convertColStringToIndex(firstCol), convertColStringToIndex(lastCol)));
//        return sheet;
//    }

    /**
     * Adjusts the width of the specified column to fit its contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets, so this should normally only be called once per column, at the
     * end of your processing.
     * 
     * @param sheet the sheet where the column is located
     * @param index the 0-based column index
     * @return the specified sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public static Sheet autoSizeColumn(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        sheet.autoSizeColumn(index);
        return sheet;
    }

    /**
     * Adjusts the width of the specified column to fit its contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets so this should normally only be called once per column at the end
     * of your processing.
     * 
     * @param sheet  the sheet where the column is located
     * @param colref the letter reference of the column
     * @return the specified sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public static Sheet autoSizeColumn(final Sheet sheet, final String colref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(colref, "colref == null");
        sheet.autoSizeColumn(convertColStringToIndex(colref));
        return sheet;
    }

//    /**
//     * Adjusts the width of all columns to fit their contents.
//     * <p>
//     * The content of merged cells is ignored.
//     * <p>
//     * This process can be relatively slow on large sheets, so this should normally only be called once per column, at the
//     * end of your processing.
//     * 
//     * @param sheet the sheet where the column is located
//     * @return the specified sheet
//     * @see Sheet#autoSizeColumn(int, boolean)
//     */
//    public static Sheet autoSizeColumns(final Sheet sheet) {
//        checkNotNull(sheet, "sheet == null");
//
//        final short max = Streams.stream(sheet).map(Row::getLastCellNum).max(Comparator.naturalOrder()).orElse((short) 0);
//
//        for (int index = 0; index < max; index++)
//            sheet.autoSizeColumn(index);
//        return sheet;
//    }

    /**
     * Adjusts the width of all columns to fit their contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets, so this should normally only be called once per column, at the
     * end of your processing.
     * 
     * @param sheet the sheet where the column is located
     * @return the specified sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public static Sheet autoSizeColumns(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");

        final short max = Streams.stream(sheet).map(Row::getLastCellNum).max(Comparator.naturalOrder()).orElse((short) 0);

        for (int index = 0; index < max; index++) {
            sheet.autoSizeColumn(index);
            final int width = sheet.getColumnWidth(index);
            sheet.setColumnWidth(index, Math.min(width + DEFAULT_PADDING, MAX_COLUMN_WIDTH));
        }

        return sheet;
    }

    private static final int DEFAULT_PADDING  = 640;
    private static final int MAX_COLUMN_WIDTH = 255 * 256;

    /**
     * Clones a sheet.
     * 
     * @param sheet the sheet to clone
     * @param name  the name of the target sheet
     * @return the target sheet
     */
    public static Sheet clone(final Sheet sheet, final String name) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final Workbook workbook = sheet.getWorkbook();
        final Sheet    target   = workbook.cloneSheet(workbook.getSheetIndex(sheet));
        return Sheets.setSheetName(target, name);
    }

    /**
     * Returns the specified row or {@code null} if it does not exist.
     * 
     * @param sheet the sheet where the row is located
     * @param index the 0-based row index
     * @return the specified row or {@code null} if it does not exist
     */
    public static Row getRow(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        return sheet.getRow(index);
    }

    /**
     * Returns the specified row. If the row does not exist it will be created.
     * 
     * @param sheet the sheet where the row is located
     * @param index the 0-based row index
     * @return the specified row
     */
    public static Row getOrCreateRow(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        return CellUtil.getRow(index, sheet);
    }

    /**
     * Returns the workbook that contains the specified sheet. If the sheet has been deleted this method will result in an
     * exception.
     * 
     * @param sheet the specified sheet
     * @return the workbook which contains the specified sheet
     */
    public static Workbook getWorkbookOf(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        return sheet.getWorkbook();
    }

    /**
     * Makes a column invisible.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @return the specified sheet
     */
    public static Sheet hideColumn(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        sheet.setColumnHidden(index, true);
        return sheet;
    }

    /**
     * Makes a column invisible.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @return the specified sheet
     */
    public static Sheet hideColumn(final Sheet sheet, final String colref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(colref, "colref == null");
        sheet.setColumnHidden(convertColStringToIndex(colref), true);
        return sheet;
    }

    /**
     * Inserts a row at the specified location shifting all subsequent rows by 1.
     * 
     * @param sheet the sheet in which the row will be inserted
     * @param index the 0-based row index
     * @return the new row
     */
    public static Row insertRow(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        sheet.shiftRows(index, sheet.getLastRowNum(), 1);
        return sheet.createRow(index);
    }

    /**
     * Applies a cell-style to future and existing cells in the specified column.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @param style the cell-style to set
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final int index, final CellStyle style) {
        return setColumnStyle(sheet, index, style, true);
    }

    /**
     * Applies a cell-style to future and existing cells in the specified column.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @param style the cell-style to set
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final String colref, final CellStyle style) {
        return setColumnStyle(sheet, colref, style, true);
    }

    /**
     * Applies a cell-style to the specified column.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @param style  the cell-style to set
     * @param update whether or not to update the style of existing cells
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final int index, final CellStyle style, final boolean update) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        checkNotNull(style, "style == null");

        if (update)
            for (final Row row : sheet) {
                final Cell cell = Rows.getCell(row, index);
                if (cell != null)
                    cell.setCellStyle(style);
            }

        sheet.setDefaultColumnStyle(index, style);

        return sheet;
    }

    /**
     * Applies a cell-style to the specified column.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @param style  the cell-style to set
     * @param update whether or not to update the style of existing cells
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final String colref, final CellStyle style, final boolean update) {
        checkNotNull(colref, "colref == null");
        return setColumnStyle(sheet, convertColStringToIndex(colref), style, update);
    }

    /**
     * Returns the column style for the given column or {@code null} if no style is set.
     * <p>
     * <b>Note:</b> While the API specification for {@link Sheet#getColumnStyle(int)} dictates returning {@code null} if no
     * column style is set, some implementations incorrectly return the default workbook style. This method explicitly
     * checks if the retrieved style is the default workbook style, and in such case returns {@code null}.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @return the column style for the given column or {@code null}
     */
    public static CellStyle getColumnStyle(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");

        final CellStyle style = sheet.getColumnStyle(index);

        if (sheet instanceof HSSFSheet)
            return style;
        else if (sheet instanceof XSSFSheet || sheet instanceof SXSSFSheet) {
            final XSSFSheet    xssfSheet = sheet instanceof XSSFSheet ? (XSSFSheet) sheet : getXSSFSheet((SXSSFSheet) sheet);
            final ColumnHelper helper    = xssfSheet.getColumnHelper();
            return sheet.getWorkbook().getCellStyleAt(helper.getColDefaultStyle(index));
        } else
            throw new IllegalArgumentException("unsupported sheet class: " + sheet.getClass().getSimpleName());
    }

    /**
     * Returns the column style for the given column or {@code null} if no style is set.
     * <p>
     * <b>Note:</b> While the API specification for {@link Sheet#getColumnStyle(int)} dictates returning {@code null} if no
     * column style is set, some implementations incorrectly return the default workbook style instead. This method
     * explicitly checks if the retrieved style is the default workbook style, and in such case returns {@code null}.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @return tthe column style for the given column or {@code null}
     */
    public static CellStyle getColumnStyle(final Sheet sheet, final String colref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(colref, "colref == null");
        return getColumnStyle(sheet, convertColStringToIndex(colref));
    }

    /**
     * Sets the width of a column in units of roughly 1 character width.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @param width the width of the column in units of roughly 1 character width
     * @return the specified sheet
     */
    public static Sheet setColumnWidth(final Sheet sheet, final int index, final int width) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        checkArgument(width > 0, "width < 0");
        sheet.setColumnWidth(index, width * 256);
        return sheet;
    }

    /**
     * Sets the width of a column in units of roughly 1 character width.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @param width  the width of the column in units of roughly 1 character width
     * @return the specified sheet
     */
    public static Sheet setColumnWidth(final Sheet sheet, final String colref, final int width) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(colref, "colref == null");
        checkArgument(width > 0, "width <= 0");
        sheet.setColumnWidth(convertColStringToIndex(colref), width * 256);
        return sheet;
    }

    /**
     * Sets the height for future and existing rows in the specified sheet.
     * 
     * @param sheet  the specified sheet
     * @param height the height to set in points
     * @return the specified sheet
     */
    public static Sheet setRowHeight(final Sheet sheet, final float height) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(height > 0, "height < 1");

        for (final Row row : sheet)
            row.setHeightInPoints(height);

        sheet.setDefaultRowHeightInPoints(height);

        return sheet;
    }

    /**
     * Sets the name of the specified sheet.
     * 
     * @param sheet the specified sheet
     * @param name  the name to set
     * @return the specified sheet
     * @throws IllegalArgumentException if the name contains illegal characters
     */
    public static Sheet setSheetName(final Sheet sheet, final String name) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final Workbook workbook = sheet.getWorkbook();
        final int      index    = workbook.getSheetIndex(sheet);
        workbook.setSheetName(index, name);
        return sheet;
    }

//    static String validateSheetName(final String name) {
//        checkNotNull(name, "name == null");
//
//        checkArgument(name.length() > 0 && name.length() < 32, "name.length() must be between 1 and 31 characters");
//
//        for (final char ch : new char[] { '/', '\\', '?', '*', ']', '[', ':' })
//            checkArgument(name.indexOf(ch) < 0, "found invalid character: %s", ch);
//
//        return name;
//    }

//    /**
//     * Sets the zoom magnification for the specified sheet.
//     * 
//     * @param sheet   the specified sheet
//     * @param percent the zoom percentage in integer units
//     * @return the specified sheet
//     */
//    public static Sheet setZoom(final Sheet sheet, final int percent) {
//        checkNotNull(sheet, "sheet == null");
//        checkArgument(percent >= 0 && percent <= 200, "percent must be between 0 and 200 inclusive");
//        sheet.setZoom(percent, 100);
//        return sheet;
//    }
//
//    /**
//     * Returns a view of the specified sheet skipping blank rows.
//     * <p>
//     * A row is considered <i>blank</i> if the {@link Cells#formatValue(Cell)} method returns an empty {@code String}, or a
//     * {@code String} composed of only whitespace characters, according to {@link CharMatcher#WHITESPACE} for every cell in
//     * the row.
//     * 
//     * @param sheet the specified sheet
//     * @return a view of the specified sheet skipping blank rows
//     */
//    public static Iterable<Row> skipBlankRows(final Iterable<Row> sheet) {
//        checkNotNull(sheet, "sheet == null");
//        return Iterables.filter(sheet, new Predicate<Row>() {
//
//            @Override
//            public boolean apply(Row row) {
//                for (final Cell cell : row)
//                    if (!isWhitespace(Cells.formatValue(cell)))
//                        return true;
//                return false;
//            }
//
//        });
//    }
//
//    /**
//     * Returns a view of the specified sheet skipping the first row.
//     * 
//     * @param sheet the specified sheet
//     * @return a view of the specified sheet that skips the first row
//     */
//    public static Iterable<Row> skipFirstRow(final Iterable<Row> sheet) {
//        checkNotNull(sheet, "sheet == null");
//        return Iterables.skip(sheet, 1);
//    }
//
//    /**
//     * Returns a view of the specified sheet starting with the given row.
//     * 
//     * @param sheet the specified sheet
//     * @param n     the 0-based index of the first row to return
//     * @return a view of the specified sheet starting with the given row
//     */
//    public static Iterable<Row> startAtRow(final Iterable<Row> sheet, final int n) {
//        checkNotNull(sheet, "sheet == null");
//        checkArgument(n >= 0, "n < 0");
//        return Iterables.skip(sheet, n);
//    }

    /**
     * Makes a column visible.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @return the specified sheet
     */
    public static Sheet unhideColumn(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");
        sheet.setColumnHidden(index, false);
        return sheet;
    }

    /**
     * Makes a column visible.
     * 
     * @param sheet  the specified sheet
     * @param colref the letter reference of the column
     * @return the specified sheet
     */
    public static Sheet unhideColumn(final Sheet sheet, final String colref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(colref, "colref == null");
        sheet.setColumnHidden(convertColStringToIndex(colref), false);
        return sheet;
    }

    /**
     * Creates and returns the next available row in the specified sheet. Shorthand for
     * {@code sheet.createRow(sheet.getLastRowNum() + 1)}.
     * 
     * @param sheet the specified sheet
     * @return the next available row
     */
    public static Row createNextRow(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        return sheet.createRow(sheet.getLastRowNum() + 1); // does this work for row 0?
    }

//    /**
//     * Enables filtering for the range of cells covering the entire sheet.
//     * 
//     * @param sheet the specified sheet
//     * @return the specified sheet
//     */
//    public static Sheet autoFilterTopRow(final Sheet sheet) {
//        checkNotNull(sheet, "sheet == null");
//
//        final Row first = sheet.getRow(0);
//
//        if (first != null) {
//            int lastCellNum = 0;
//            for (final Row row : sheet)
//                lastCellNum = Math.max(lastCellNum, row.getLastCellNum());
//
//            final Row last = getLastRow(sheet);
//
//            Sheets.autoFilter(sheet, first, 0, last, lastCellNum - 1);
//        }
//        return sheet;
//    }

    /**
     * Returns the last row in the specified sheet or {@code null} if it doesn't exist.
     * 
     * @param sheet the specified sheet
     * @return the last row in the specified sheet or {@code null} if it doesn't exist
     */
    public static Row getLastRow(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        return sheet.getRow(sheet.getLastRowNum());
    }

    /**
     * Creates a freeze pane that keeps the top row visible while scrolling through the specified sheet.
     * 
     * @param sheet the specified sheet
     * @return the specified sheet
     */
    public static Sheet freezeTopRow(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        sheet.createFreezePane(0, 1);
        return sheet;
    }

    /**
     * Removes any existing freeze pane from the sheet.
     * 
     * @param sheet the specified sheet
     * @return the specified sheet
     */
    public static Sheet removeFreezePane(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        sheet.createFreezePane(0, 0);
        return sheet;
    }

    /**
     * Removes a row from the specified sheet.
     * 
     * @param sheet the sheet to remove the row from
     * @param index the 0-based row index
     * @return the specified sheet
     */
    public static Sheet removeRow(final Sheet sheet, final int index) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "index < 0");

        final int last = sheet.getLastRowNum();

        if (last != -1) {
            if (index >= 0 && index < last)
                sheet.shiftRows(index + 1, last, -1);

            sheet.removeRow(sheet.getRow(last));
        }

        return sheet;
    }

    private static XSSFSheet getXSSFSheet(final SXSSFSheet sheet) {
        try {
            final Field field = sheet.getClass().getDeclaredField("_sh");
            field.setAccessible(true);
            return (XSSFSheet) field.get(sheet);
        } catch (final NoSuchFieldException | IllegalAccessException e) {
            throw new AssertionError(e);
        }
    }

}