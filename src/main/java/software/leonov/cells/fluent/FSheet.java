package software.leonov.cells.fluent;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.util.Comparator;
import java.util.Iterator;
import java.util.concurrent.ExecutionException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;

import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.google.common.collect.Iterators;
import com.google.common.collect.Streams;

/**
 * A representation of a sheet in a Microsoft Excel workbook.
 * 
 * @author Zhenya Leonov
 */
public final class FSheet implements Iterable<FRow> {

    private static final Cache<Row, FRow> rows = CacheBuilder.newBuilder().maximumSize(1000).build();

    private final FWorkbook fworkbook;
    private final Sheet sheet;

    FSheet(final FWorkbook fworkbook, final Sheet sheet) {
        checkNotNull(fworkbook, "fworkbook == null");
        checkNotNull(sheet, "sheet == null");
        this.fworkbook = fworkbook;
        this.sheet = sheet;
    }

    Sheet delegate() {
        return sheet;
    }

    /**
     * Copies the cell-style, cell-type, comment, and value from the source cell to the target cell. If the target cell
     * contains a value it will be overwritten.
     * <p>
     * Note: Both cells must be located in the same workbook.
     * 
     * @param from the source cell
     * @param to   the target cell
     * @return the target cell
     */
    public static FCell copyCell(final FCell from, final FCell to) {
        checkNotNull(from, "from == null");
        checkNotNull(to, "to == null");
        checkArgument(from.getRow().getSheet().getWorkbook().delegate().equals(to.getRow().getSheet().delegate().getWorkbook()), "the source cell is not located in the same workbook as the target cell");

        to.setStyle(from.getStyle());
        to.setComment(from.getComment());

        switch (from.getType()) {
        case NUMERIC:
            to.delegate().setCellValue(from.delegate().getNumericCellValue());
            break;
        case STRING:
            to.delegate().setCellValue(from.delegate().getStringCellValue());
            break;
        case FORMULA:
            to.delegate().setCellValue(from.delegate().getCellFormula());
            break;
        case BOOLEAN:
            to.delegate().setCellValue(from.delegate().getBooleanCellValue());
            break;
        case ERROR:
            to.delegate().setCellValue(from.delegate().getBooleanCellValue());
            break;
        case BLANK:
            to.delegate().setCellValue((String) null);
        default: // examine _NONE style?
            break;
        }
        return to;
    }

    /**
     * Cuts and pastes the cell-style, cell-type, comment, and value from the source cell to the target cell. If the target
     * cell contains a value it will be overwritten.
     * <p>
     * Note: Both cells must be located in the same workbook.
     * 
     * @param from the source cell
     * @param to   the target cell
     * @return the target cell
     */
    public static FCell cutAndPasteCell(final FCell from, final FCell to) {
        copyCell(from, to);
        from.getRow().removeCell(from);
        return to;
    }

    /**
     * Enables filtering for the given range of cells in this sheet.
     * 
     * @param from  the first row
     * @param start the 0-based index of the first column
     * @param to    the last row
     * @param end   the 0-based index of the last column
     * @return this sheet
     */
    public FSheet autoFilter(final FRow from, final int start, final FRow to, final int end) {
        checkNotNull(from, "from == null");
        checkNotNull(to, "to == null");
        sheet.setAutoFilter(new CellRangeAddress(from.delegate().getRowNum(), to.delegate().getRowNum(), start, end));
        return this;
    }

    /**
     * Enables filtering for the given range of cells in this sheet.
     * 
     * @param from  the first row
     * @param start the letter reference of first column
     * @param to    the last row
     * @param end   the letter reference of the last column
     * @return this sheet
     */
    public FSheet autoFilter(final FRow from, final String start, final FRow to, final String end) {
        checkNotNull(from, "from == null");
        checkNotNull(start, "start == null");
        checkNotNull(to, "to == null");
        checkNotNull(end, "end == null");
        sheet.setAutoFilter(new CellRangeAddress(from.delegate().getRowNum(), to.delegate().getRowNum(), CellReference.convertColStringToIndex(start), CellReference.convertColStringToIndex(end)));
        return this;
    }

    /**
     * Adjusts the width of the specified column to fit its contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets, so this should normally only be called once per column, at the
     * end of your processing.
     * 
     * @param index the 0-based column index
     * @return this sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public FSheet autoSizeColumn(final int index) {
        checkArgument(index >= 0, "index < 0");
        sheet.autoSizeColumn(index);
        return this;
    }

    /**
     * Adjusts the width of the specified column to fit its contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets so this should normally only be called once per column at the end
     * of your processing.
     * 
     * @param ref the letter reference of the column
     * @return this sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public FSheet autoSizeColumn(final String ref) {
        checkNotNull(ref, "ref == null");
        sheet.autoSizeColumn(CellReference.convertColStringToIndex(ref));
        return this;
    }

    /**
     * Adjusts the width of all columns to fit their contents.
     * <p>
     * The content of merged cells is ignored.
     * <p>
     * This process can be relatively slow on large sheets, so this should normally only be called once per column, at the
     * end of your processing.
     * 
     * @return this sheet
     * @see Sheet#autoSizeColumn(int, boolean)
     */
    public FSheet autoSizeColumns() {
        short max = Streams.stream(sheet).map(Row::getLastCellNum).max(Comparator.naturalOrder()).orElse((short) 0);

        for (int index = 0; index <= max; index++)
            sheet.autoSizeColumn(index);
        return this;
    }

    /**
     * Creates and returns the next available row in this sheet.
     * 
     * @return the next available row
     */
    public FRow createNextRow() {
        return getOrCreateFromCache(sheet.createRow(sheet.getLastRowNum() + 1)); // does this work for row 0?
    }

    /**
     * Returns the specified row. If the row does not exist it will be created.
     * 
     * @param rownum the 0-based index of the specified row
     * @return the specified row
     */
    public FRow getOrCreateRow(final int rownum) {
        checkArgument(rownum >= 0, "rownum < 0");
        return getOrCreateFromCache(CellUtil.getRow(rownum, sheet));
    }
    
//    public FRow getRow(final int rownum) {
//        checkArgument(rownum >= 0, "rownum < 0");
//        final Row row = sheet.getRow(rownum);
//        return row == null ? null : getOrCreateFromCache(row);
//    }

    /**
     * Returns the workbook that contains this sheet.
     * 
     * @return the workbook which contains this sheet
     */
    public FWorkbook getWorkbook() {
        return fworkbook;
    }

    /**
     * Makes a column invisible.
     * 
     * @param sheet the specified sheet
     * @param index the 0-based column index
     * @return this sheet
     */
    public FSheet hideColumn(final int index) {
        checkArgument(index >= 0, "index < 0");
        sheet.setColumnHidden(index, true);
        return this;
    }

    /**
     * Makes a column invisible.
     * 
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @return the specified sheet
     */
    public FSheet hideColumn(final String ref) {
        checkNotNull(ref, "ref == null");
        sheet.setColumnHidden(CellReference.convertColStringToIndex(ref), true);
        return this;
    }

    /**
     * Inserts a row at the specified location shifting all subsequent rows by 1.
     * 
     * @param rownum the 0-based index of the row
     * @return the new row
     */
    public FRow insertRow(final int rownum) {
        checkArgument(rownum >= 0, "rownum < 0");
        sheet.shiftRows(rownum, sheet.getLastRowNum(), 1);
        return getOrCreateFromCache(sheet.createRow(rownum));
    }

    @Override
    public Iterator<FRow> iterator() {
        return Iterators.transform(sheet.iterator(), this::getOrCreateFromCache);
    }

    /**
     * Sets the column style for future and existing cells in a column.
     * 
     * @param index the 0-based column index
     * @param style the cell-style to set
     * @return this sheet
     */
    public FSheet setColumnStyle(final int index, final CellStyle style) {
        checkArgument(index >= 0, "column index < 0");
        checkNotNull(style, "style == null");

        for (final Row row : sheet) {
            final Cell cell = row.getCell(index);
            if (cell != null)
                cell.setCellStyle(style);
        }

        sheet.setDefaultColumnStyle(index, style);

        return this;
    }

    /**
     * Sets the column style for future and existing cells in a column.
     * 
     * @param ref   the letter reference of the column
     * @param style the cell-style to set
     * @return this sheet
     */
    public FSheet setColumnStyle(final String ref, final CellStyle style) {
        checkNotNull(ref, "ref == null");
        checkNotNull(style, "style == null");

        return setColumnStyle(CellReference.convertColStringToIndex(ref), style);
    }

    /**
     * Sets the width of a column in units of roughly 1 character width.
     * 
     * @param index the 0-based column index
     * @param width the width of the column in units of roughly 1 character width
     * @return this sheet
     */
    public FSheet setColumnWidth(final int index, final int width) {
        checkNotNull(index, "column == null");
        checkArgument(width > 0, "width < 0");
        sheet.setColumnWidth(index, width * 256);
        return this;
    }

    /**
     * Sets the width of a column in units of roughly 1 character width.
     * 
     * @param ref   the letter reference of the column
     * @param width the width of the column in units of roughly 1 character width
     * @return this sheet
     */
    public FSheet setColumnWidth(final String ref, final int width) {
        checkNotNull(ref, "ref == null");
        checkArgument(width > 0, "width <= 0");
        sheet.setColumnWidth(CellReference.convertColStringToIndex(ref), width * 256);
        return this;
    }

    /**
     * Sets the height for future and existing rows in this sheet.
     * 
     * @param height the height to set in points
     * @return this sheet
     */
    public FSheet setRowHeight(final float height) {
        checkArgument(height > 0, "height < 1");

        for (final Row row : sheet)
            row.setHeightInPoints(height);

        sheet.setDefaultRowHeightInPoints(height);

        return this;
    }

    /**
     * Sets the name of this sheet.
     * 
     * @param name the name to set
     * @return this workbook
     * @throws IllegalArgumentException if the name contains illegal characters
     */
    public FSheet setSheetName(final String name) {
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final Workbook workbook = sheet.getWorkbook();
        final int index = workbook.getSheetIndex(sheet);
        workbook.setSheetName(index, name);
        return this;
    }

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
     * Sets the zoom magnification for this sheet.
     * 
     * @param percent the zoom percentage in integer units from 10 to 400
     * @return this sheet
     */
    public FSheet setZoom(final int percent) {
        checkArgument(percent >= 10 && percent <= 400, "percent must be between 10 and 400 inclusive");
        sheet.setZoom(percent);
        return this;
    }

    /**
     * Makes a column visible.
     * 
     * @param index the 0-based column index
     * @return this sheet
     */
    public FSheet unhideColumn(final int index) {
        checkArgument(index >= 0, "column index < 0");
        sheet.setColumnHidden(index, false);
        return this;
    }

    /**
     * Makes a column visible.
     * 
     * @param sheet the specified sheet
     * @param ref   the letter reference of the column
     * @return this sheet
     */
    public FSheet unhideColumn(final String ref) {
        checkNotNull(ref, "ref == null");
        sheet.setColumnHidden(CellReference.convertColStringToIndex(ref), false);
        return this;
    }

    private FRow getOrCreateFromCache(final Row row) {
        try {
            return rows.get(row, () -> new FRow(this, row));
        } catch (ExecutionException e) {
            throw new AssertionError(); // cannot happen
        }
    }

    public FSheet createFreezePane(int i, int j) {
        sheet.createFreezePane(i, j);
        return this;        
    }

}