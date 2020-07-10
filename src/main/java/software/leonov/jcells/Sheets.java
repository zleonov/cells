package software.leonov.jcells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.util.Comparator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;

import com.google.common.collect.Streams;

/**
 * Static methods for working with {@link Sheet}s.
 * 
 * @author Zhenya Leonov
 */
final public class Sheets {

    private Sheets() {
    }

    /**
     * Enables filtering for the given range of cells in the specified sheet.
     * 
     * @param sheet      the specified sheet
     * @param fromRow    the first row
     * @param fromColumn the 0-based index of the first column
     * @param toRow      the last row
     * @param toColumn   the 0-based index of the last column
     * @return the specified sheet
     */
    public static Sheet autoFilter(final Sheet sheet, final Row fromRow, final int fromColumn, final Row toRow, final int toColumn) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(fromRow, "fromRow == null");
        checkNotNull(toRow, "toRow == null");
        sheet.setAutoFilter(new CellRangeAddress(fromRow.getRowNum(), toRow.getRowNum(), fromColumn, toColumn));
        return sheet;
    }

    /**
     * Enables filtering for the given range of cells in the specified sheet.
     * 
     * @param sheet      the specified sheet
     * @param fromRow    the first row
     * @param fromColumn the letter reference of first column
     * @param toRow      the last row
     * @param toColumn   the letter reference of the last column
     * @return the specified sheet
     */
    public static Sheet autoFilter(final Sheet sheet, final Row fromRow, final String fromColumn, final Row toRow, final String toColumn) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(fromRow, "fromRow == null");
        checkNotNull(fromColumn, "fromColumn == null");
        checkNotNull(toRow, "toRow == null");
        checkNotNull(toColumn, "toColumn == null");
        sheet.setAutoFilter(new CellRangeAddress(fromRow.getRowNum(), toRow.getRowNum(), Columns.getIndex(fromColumn), Columns.getIndex(toColumn)));
        return sheet;
    }

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
     * @param sheet the sheet where the column is located
     * @param ref   the letter reference of the column
     * @return the specified sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public static Sheet autoSizeColumn(final Sheet sheet, final String ref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(ref, "ref == null");
        sheet.autoSizeColumn(Columns.getIndex(ref));
        return sheet;
    }

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

        short max = Streams.stream(sheet).map(Row::getLastCellNum).max(Comparator.naturalOrder()).orElse((short) 0);

        for (int index = 0; index <= max; index++)
            sheet.autoSizeColumn(index);
        return sheet;
    }

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
        final Sheet target = workbook.cloneSheet(workbook.getSheetIndex(sheet));
        return Sheets.setSheetName(target, name);
    }

    /**
     * Returns the specified row. If the row does not exist it will be created.
     * 
     * @param sheet  the sheet where the row is located
     * @param rownum the 0-based index of the specified row
     * @return the specified row
     */
    public static Row getOrCreateRow(final Sheet sheet, final int rownum) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(rownum >= 0, "rownum < 0");
        return CellUtil.getRow(rownum, sheet);
    }

    /**
     * Returns the workbook that contains the specified sheet. If the sheet has been deleted this method will result in an
     * exception.
     * 
     * @param sheet the specified sheet
     * @return the workbook which contains the specified row
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
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @return the specified sheet
     */
    public static Sheet hideColumn(final Sheet sheet, final String ref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(ref, "ref == null");
        sheet.setColumnHidden(Columns.getIndex(ref), true);
        return sheet;
    }

    /**
     * Inserts a row at the specified location shifting all subsequent rows by 1.
     * 
     * @param sheet  the sheet in which the row will be inserted
     * @param rownum the 0-based index of the row
     */
    public static Row insertRow(final Sheet sheet, final int rownum) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(rownum >= 0, "rownum < 0");
        sheet.shiftRows(rownum, sheet.getLastRowNum(), 1);
        return sheet.createRow(rownum);
    }

    /**
     * Sets the column style for future and existing cells in the column.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @param style the cell-style to set
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final int index, final CellStyle style) {
        checkNotNull(sheet, "sheet == null");
        checkArgument(index >= 0, "column index < 0");
        checkNotNull(style, "style == null");

        for (final Row row : sheet) {
            final Cell cell = Rows.getCell(row, index);
            if (cell != null)
                cell.setCellStyle(style);
        }

        sheet.setDefaultColumnStyle(index, style);

        return sheet;
    }

    /**
     * Sets the column style future and existing cells in the column.
     * 
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @param style the cell-style to set
     * @return the specified sheet
     */
    public static Sheet setColumnStyle(final Sheet sheet, final String ref, final CellStyle style) {
        checkNotNull(sheet, "rows == null");
        checkNotNull(ref, "ref == null");
        checkNotNull(style, "style == null");

        return setColumnStyle(sheet, Columns.getIndex(ref), style);
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
        checkNotNull(index, "column == null");
        checkArgument(width > 0, "width < 0");
        sheet.setColumnWidth(index, width * 256);
        return sheet;
    }

    /**
     * Sets the width of a column in units of roughly 1 character width.
     * 
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @param width the width of the column in units of roughly 1 character width
     * @return the specified sheet
     */
    public static Sheet setColumnWidth(final Sheet sheet, final String ref, final int width) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(ref, "ref == null");
        checkArgument(width > 0, "width <= 0");
        sheet.setColumnWidth(Columns.getIndex(ref), width * 256);
        return sheet;
    }

    /**
     * Sets the height for future and existing rows in the specified sheet.
     * 
     * @param sheet  the specified sheet
     * @param height the height to set in points
     * @return the specified sheet
     */
    public static Iterable<Row> setRowHeight(final Sheet sheet, final float height) {
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
        final int index = workbook.getSheetIndex(sheet);
        workbook.setSheetName(index, name);
        return sheet;
    }

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
        checkArgument(index >= 0, "column index < 0");
        sheet.setColumnHidden(index, false);
        return sheet;
    }

    /**
     * Makes a column visible.
     * 
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @return the specified sheet
     */
    public static Sheet unhideColumn(final Sheet sheet, final String ref) {
        checkNotNull(sheet, "sheet == null");
        checkNotNull(ref, "ref == null");
        sheet.setColumnHidden(Columns.getIndex(ref), false);
        return sheet;
    }

}