package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;
import static software.leonov.common.base.Str.isWhitespace;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.CharMatcher;
import com.google.common.collect.Iterables;
import com.google.common.collect.Streams;

import software.leonov.common.base.Obj;

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
     * @param row    the row where the cell is located
     * @param column the 0-based column number
     * @return the specified cell or {@code null}
     */
    public static Cell getCell(final Row row, final int column) {
        checkNotNull(row, "row == null");
        checkArgument(column >= 0, "column index < 0");
        return row.getCell(column);
    }

    /**
     * Returns the specified cell or {@code null} if the cell is undefined.
     * 
     * @param row the row where the cell is located
     * @param ref the letter reference of the column
     * @return the specified cell or {@code null}
     */
    public static Cell getCell(final Row row, final String ref) {
        checkNotNull(row, "row == null");
        checkNotNull(ref, "ref == null");
        return row.getCell(Columns.get(ref));
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * <p>
     * The cell-style will be inherited from the default column style. If the default column style is undefined the
     * cell-style will be inherited from the default row style. If neither are defined it will have a {@link CellType#BLANK}
     * style.
     * 
     * @param row   the row where the cell is located
     * @param index the 0-based column index
     * @see Sheet#getColumnStyle(int)
     * @see Row#getRowStyle()
     * @return the specified cell
     */
    public static Cell getOrCreateCell(final Row row, final int index) {
        checkNotNull(row, "row == null");
        checkArgument(index >= 0, "index < 0");

        Cell cell = row.getCell(index);

        if (cell == null) {
            cell = row.createCell(index);
            final CellStyle style = Obj.coalesce(row.getSheet().getColumnStyle(index), row.getRowStyle());
            if (style != null)
                cell.setCellStyle(style);
            row.setHeightInPoints(row.getHeightInPoints());
        }

        return cell;
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * 
     * @param row    the row where the cell is located
     * @param column the letter reference of the column
     * @return the specified cell
     */
    public static Cell getOrCreateCell(final Row row, final String ref) {
        checkNotNull(row, "row == null");
        checkNotNull(ref, "ref == null");
        final int index = Columns.get(ref);
        return getOrCreateCell(row, index);
    }

    /**
     * Returns the sheet that contains the specified row. If the row has been deleted this method will result in an
     * exception.
     * 
     * @param row the specified row
     * @return the workbook which contains the specified row
     */
    public static Sheet getSheetOf(final Row row) {
        checkNotNull(row, "row == null");
        return row.getSheet();
    }

//    /**
//     * Sets a value for a cell in the specified row.
//     * <p>
//     * Unlike {@link Cells#setValue(Cell, Object)} this method is {@code null} safe. If {@code value} is {@code null}, an
//     * empty string, or a string composed exclusively of whitespace characters according to {@link CharMatcher#WHITESPACE},
//     * this method will delete the specified cell if it exists (to insert such a string consider using
//     * {@link Cell#setCellValue(String)}).
//     * <p>
//     * If {@code value} is a {@link Number} the cell value will be set to the {@code double} value of the number by first
//     * calling {@link Number#doubleValue()} followed by {@link Cell#setCellValue(double)}.
//     * <p>
//     * If {@code value} is a {@code Boolean}, {@link Calendar}, {@link Date}, or {@link RichTextString} the cell value will
//     * be set by calling {@link Cell#setCellValue(boolean)}, {@link Cell#setCellValue(Calendar)},
//     * {@link Cell#setCellValue(Date)}, or {@link Cell#setCellValue(RichTextString)} accordingly.
//     * <p>
//     * For all other types the cell will be set to {@code value.toString()}.
//     * 
//     * @param row    the row where the cell is located
//     * @param column the letter reference of the column
//     * @param value  the value to set
//     * @return the specified row
//     * @throws IllegalArgumentException if the value to set is a string which exceeds 32767 characters
//     */
//    public static Row setCellValue(final Row row, final String column, final Object value) {
//        checkNotNull(row, "row == null");
//        checkNotNull(column, "column == null");
//        return setCellValue(row, Cells.getIndex(column), value);
//    }
//
//    /**
//     * Sets a value for a cell in the specified row.
//     * <p>
//     * Unlike {@link Cells#setValue(Cell, Object)} this method is {@code null} safe. If {@code value} is {@code null}, an
//     * empty string, or a string composed exclusively of whitespace characters according to {@link CharMatcher#WHITESPACE},
//     * this method will delete the specified cell if it exists (to insert such a string consider using
//     * {@link Cell#setCellValue(String)}).
//     * <p>
//     * If {@code value} is a {@link Number} the cell value will be set to the {@code double} value of the number by first
//     * calling {@link Number#doubleValue()} followed by {@link Cell#setCellValue(double)}.
//     * <p>
//     * If {@code value} is a {@code Boolean}, {@link Calendar}, {@link Date}, or {@link RichTextString} the cell value will
//     * be set by calling {@link Cell#setCellValue(boolean)}, {@link Cell#setCellValue(Calendar)},
//     * {@link Cell#setCellValue(Date)}, or {@link Cell#setCellValue(RichTextString)} accordingly.
//     * <p>
//     * For all other types the cell will be set to {@code value.toString()}.
//     * 
//     * @param row    the row where the cell is located
//     * @param column the 0-based column number
//     * @param value  the value to set
//     * @return the specified row
//     * @throws IllegalArgumentException if the value to set is a string which exceeds 32767 characters
//     */
//    public static Row setCellValue(final Row row, final int column, final Object value) {
//        checkNotNull(row, "row == null");
//        checkArgument(column >= 0, "column < 0");
//
//        Cell cell = row.getCell(column);
//
//        if (value == null || isWhitespace(value.toString())) {
//            if (cell != null)
//                row.removeCell(cell);
//            return row;
//        }
//
//        if (cell == null)
//            cell = getOrCreateCell(row, column);
//
//        if (value instanceof Boolean) {
//            cell.setCellValue((Boolean) value);
//            checkArgument(cell.getCellType() == CellType.BOOLEAN && Objects.equal(cell.getBooleanCellValue(), value),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        } else if (value instanceof Calendar) {
//            cell.setCellValue((Calendar) value);
//            checkArgument(cell.getCellType() == CellType.NUMERIC && Objects.equal(cell.getNumericCellValue(), ((Calendar) value).getTime()),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        } else if (value instanceof Date) {
//            cell.setCellValue((Date) value);
//            checkArgument(cell.getCellType() == CellType.NUMERIC && Objects.equal(cell.getDateCellValue(), value),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        } else if (value instanceof Number) {
//            cell.setCellValue(((Number) value).doubleValue());
//            checkArgument(cell.getCellType() == CellType.NUMERIC && Objects.equal(cell.getNumericCellValue(), ((Number) value).doubleValue()),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        } else if (value instanceof RichTextString) {
//            cell.setCellValue((RichTextString) value);
//            checkArgument(cell.getCellType() == CellType.NUMERIC && Objects.equal(cell.getRichStringCellValue(), value),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        } else {
//            final String string = value.toString();
//            checkArgument(string.length() <= 32767, "length > 32767 characters for value: %s", truncate(string, 500, "..."));
//            cell.setCellValue(string);
//            checkArgument(cell.getCellType() == Cell.CELL_TYPE_STRING && Objects.equal(cell.getStringCellValue(), string),
//                    "cannot update cell value: this error may occur when trying to update a cell in a document previously written in Format.STREAMING_OFFICE_OPEN_XML");
//        }
//
//        return row;
//    }

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
    public static Row setRowHeight(final Row row, final float height) {
        checkNotNull(row, "row == null");
        checkArgument(height > 0, "height < 1");
        row.setHeightInPoints(height);
        return row;
    }

    /**
     * Applies the cell-style to future and existing cells in the specified row.
     * 
     * @param row   the row to apply the cell-style to
     * @param style the specified cell-style
     * @return the affected row
     */
    public static Row setStyle(final Row row, final CellStyle style) {
        checkNotNull(row, "row == null");
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
        return Iterables.filter(row, cell -> !isWhitespace(Cells.formatValue(cell)));
    }
}