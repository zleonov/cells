package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;
import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;
import static software.leonov.cells.Sheets.getColumnStyle;
import static software.leonov.common.base.Obj.coalesce;

import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.CharMatcher;
import com.google.common.collect.Iterables;
import com.google.common.collect.Streams;

import software.leonov.common.base.Str;

/**
 * Static methods for working with {@link Row}s.
 * 
 * @author Zhenya Leonov
 */
public final class Rows {

    private Rows() {
    }

    /**
     * Returns the specified cell or {@code null} if the cell is undefined.
     * 
     * @param row   the row where the cell is located
     * @param index the 0-based column index
     * @return the specified cell or {@code null}
     */
    public static Cell getCell(final Row row, final int index) {
        checkNotNull(row, "row == null");
        checkArgument(index >= 0, "index < 0");
        return row.getCell(index);
    }

    /**
     * Returns the specified cell or {@code null} if the cell is undefined.
     * 
     * @param row    the row where the cell is located
     * @param colref the letter reference of the column
     * @return the specified cell or {@code null}
     */
    public static Cell getCell(final Row row, final String colref) {
        checkNotNull(row, "row == null");
        checkNotNull(colref, "colref == null");
        return row.getCell(convertColStringToIndex(colref));
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * <p>
     * The cell-style will be inherited from default styles in the following order: the row column, then the default column
     * style, and finally the default workbook style.
     * 
     * @param row   the row where the cell is located
     * @param index the 0-based column index
     * @return the specified cell
     */
    public static Cell getOrCreateCell(final Row row, final int index) {
        checkNotNull(row, "row == null");
        checkArgument(index >= 0, "index < 0");

        Cell cell = row.getCell(index);

        if (cell == null) {
            cell = row.createCell(index);
            final CellStyle style = coalesce(getRowStyle(row), getColumnStyle(getSheetOf(row), index));
            if (style != null)
                cell.setCellStyle(style);
        }

        return cell;
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * <p>
     * The cell-style will be inherited from default styles in the following order: the row column, then the default column
     * style, and finally the default workbook style.
     * 
     * @param row    the row where the cell is located
     * @param colref the letter reference of the column
     * @return the specified cell
     */
    public static Cell getOrCreateCell(final Row row, final String colref) {
        checkNotNull(row, "row == null");
        checkNotNull(colref, "colref == null");
        final int index = convertColStringToIndex(colref);
        return getOrCreateCell(row, index);
    }

    /**
     * Returns the row style for the given row or {@code null} if no style is set.
     * 
     * @param row the specified row
     * @return the row style for the given row or {@code null}
     */
    public static CellStyle getRowStyle(final Row row) {
        checkNotNull(row, "row == null");
        return row.getRowStyle();
    }

    /**
     * Returns the sheet that contains the specified row. If the row has been deleted this method will result in an
     * exception.
     * 
     * @param row the specified row
     * @return the sheet which contains the specified row
     */
    public static Sheet getSheetOf(final Row row) {
        checkNotNull(row, "row == null");
        return row.getSheet();
    }

    /**
     * Returns the workbook that contains the specified row. If the row has been deleted this method will result in an
     * exception.
     * 
     * @param row the specified row
     * @return the workbook which contains the specified row
     */
    public static Workbook getWorkbookOf(final Row row) {
        checkNotNull(row, "row == null");
        return Sheets.getWorkbookOf(getSheetOf(row));
    }

    /**
     * Sets the height of the specified row.
     * 
     * @param row    the specified row
     * @param height the height to set, in points
     * @return the specified row
     */
    public static Row setHeight(final Row row, final float height) {
        checkNotNull(row, "row == null");
        checkArgument(height > 0, "height < 1");
        row.setHeightInPoints(height);
        return row;
    }

    /**
     * Applies a cell-style to future and existing cells in the specified row.
     * 
     * @param row   the row to apply the cell-style to
     * @param style the specified cell-style
     * @return the affected row
     */
    public static Row setStyle(final Row row, final CellStyle style) {
        return setStyle(row, style, true);
    }

    /**
     * Applies a cell-style to the specified row.
     * 
     * @param row    the row to apply the cell-style to
     * @param style  the specified cell-style
     * @param update whether or not to update existing cells
     * @return the affected row
     */
    public static Row setStyle(final Row row, final CellStyle style, final boolean update) {
        checkNotNull(row, "row == null");
        checkNotNull(style, "style == null");
        if (update)
            Streams.stream(row).forEach(cell -> Cells.setStyle(cell, style));
        row.setRowStyle(style);
        return row;
    }

    /**
     * Returns a view of the specified row skipping blank cells.
     * <p>
     * A cell is considered <i>blank</i> if the {@link Cells#formatValue(Cell)} method returns an empty {@code String}, or a
     * {@code String} composed of only whitespace characters, according to {@link CharMatcher#whitespace()}.
     * 
     * @param row the specified row
     * @return a view of the specified row skipping blank cells
     */
    public static Iterable<Cell> skipBlankCells(final Row row) {
        checkNotNull(row, "row == null");
        return Iterables.filter(row, cell -> Cells.formatValue(cell) != null);
    }

    /**
     * Returns the index (0-based) of the first cell in the specified row or an empty {@code Optional} if the row has no
     * defined cells.
     * 
     * @param row the specified row
     * @return index (0-based) of the first cell in the specified row or an empty {@code Optional} if the row has no defined
     *         cells
     */
    public static Optional<Integer> getFirstCellIndex(final Row row) {
        checkNotNull(row, "row == null");
        final int i = row.getFirstCellNum();
        return i < 0 ? Optional.empty() : Optional.of(i);
    }

    /**
     * Returns the index (0-based) of the last cell in the specified row or an empty {@code Optional} if the row has no
     * defined cells.
     * 
     * @param row the specified row
     * @return index (0-based) of the last cell in the specified row or an empty {@code Optional} if the row has no defined
     *         cells
     */
    public static Optional<Integer> getLastCellIndex(final Row row) {
        checkNotNull(row, "row == null");
        final int i = row.getLastCellNum();
        return i < 0 ? Optional.empty() : Optional.of(i - 1);
    }

    /**
     * Sets a sequence of values in the given row, beginning at the specified cell.
     *
     * Any non-existent cells within the range are created. The values are set by calling
     * {@link Cells#setValue(Cell, Object)}.
     * 
     * @param row    the specified row
     * @param index  the 0-based index of the starting cell
     * @param values the values to set
     * @return the specified row
     */
    public static Row setValues(final Row row, int index, final Iterable<? extends Object> values) {
        checkNotNull(row, "row == null");
        checkNotNull(values, "values == null");
        checkArgument(index >= 0, "index < 0");

        final Iterator<? extends Object> itor = values.iterator();

        for (final Object value : values) {
            final Cell cell = getCell(row, index);
            if (itor.hasNext() && cell != null && isBlank(value))
                row.removeCell(cell);
            else
                Cells.setValue(getOrCreateCell(row, index++), value);
        }

        return row;
    }

    private static boolean isBlank(final Object value) {
        if (value == null)
            return true;
        if (value instanceof Boolean || value instanceof Calendar || value instanceof Date || value instanceof Number || value instanceof LocalDateTime || value instanceof RichTextString)
            return false;
        else
            return Str.isWhitespace(value.toString());
    }

    /**
     * Sets a sequence of values in the given row, beginning at the specified cell.
     *
     * Any non-existent cells within the range are created. The values are set by calling
     * {@link Cells#setValue(Cell, Object)}.
     * 
     * @param row    the specified row
     * @param colref the letter reference of the starting cell
     * @param values the values to set
     * @return the specified row
     */
    public static Row setValues(final Row row, final String colref, final Iterable<? extends Object> values) {
        checkNotNull(colref, "colref == null");
        return setValues(row, convertColStringToIndex(colref), values);
    }

    /**
     * Creates and returns the next available cell in the specified row.
     * 
     * @param row the specified row
     * @return the next available cell
     */
    public static Cell createNextCell(final Row row) {
        checkNotNull(row, "row == null");
        final int idx = row.getLastCellNum();
        return idx == -1 ? getOrCreateCell(row, 0) : getOrCreateCell(row, idx);
    }

//    /**
//     * Returns the index (0-based) of the first cell in the specified row.
//     * 
//     * @param row the specified row
//     * @return the index (0-based) of the first cell in the specified row
//     */
//    public static int getFirstCellIndex(final Row row) {
//        checkNotNull(row, "row == null");
//        return row.getFirstCellNum();
//    }

//    /**
//     * Returns the index (0-based) of the last cell in the specified row.
//     * 
//     * @param row the specified row
//     * @return the index (0-based) of the last cell in the specified row
//     */
//    public static int getLastCellIndex(final Row row) {
//        checkNotNull(row, "row == null");
//        return row.getLastCellNum() - 1;
//    }

}