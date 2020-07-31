package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;
import static software.leonov.common.base.Str.isWhitespace;
import static software.leonov.common.base.Str.trim;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.CharMatcher;

/**
 * Static methods for working with {@link Cell}s.
 * <p>
 * Some methods in this class are specifically documented as being {@code null} safe. Which means they will return
 * {@code null} values (unless otherwise stated) instead of throwing {@code NullPointerException}s when given
 * {@code null} arguments. All other methods should be expected to throw {@code Exception}s in the presence of
 * {@code null} inputs.
 * 
 * @author Zhenya Leonov
 */
final public class Cells {

    /**
     * The total number of characters that a cell can contain
     */
    public static final int MAX_CELL_SIZE = 32767;

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    private Cells() {
    }

    /**
     * Copies the cell-style, cell-type, comment, and value from the specified cell to the target cell. If the target cell
     * contains a value it will be overwritten.
     * <p>
     * Note: Both cells must be located in the same workbook.
     * 
     * @param from the specified cell
     * @param to   the target cell
     * @return the target cell
     */
    public static Cell copy(final Cell from, final Cell to) {
        checkNotNull(from, "from == null");
        checkNotNull(to, "to == null");
        checkArgument(getWorkbookOf(from).equals(getWorkbookOf(to)), "the specified Cell is not located in the same Workbook as the target Cell");

        to.setCellStyle(from.getCellStyle());
        to.setCellComment(from.getCellComment());

        switch (from.getCellType()) {
        case NUMERIC:
            to.setCellValue(from.getNumericCellValue());
            break;
        case STRING:
            to.setCellValue(from.getStringCellValue());
            break;
        case FORMULA:
            to.setCellValue(from.getCellFormula());
            break;
        case BOOLEAN:
            to.setCellValue(from.getBooleanCellValue());
            break;
        case ERROR:
            to.setCellValue(from.getBooleanCellValue());
            break;
        case BLANK:
            to.setCellValue((String) null);
        default: // examine _NONE style?
            break;
        }
        return to;
    }

    /**
     * Cuts and pastes the cell-style, cell-type, comment, and value from the specified cell to the target cell. If the
     * target cell contains a value it will be overwritten.
     * <p>
     * Note: Both cells must be located in the same workbook.
     * 
     * @param from the source cell
     * @param to   the target cell
     * @return the target cell
     */
    public static Cell cutAndPaste(final Cell from, final Cell to) {
        checkNotNull(from, "from == null");
        checkNotNull(to, "to == null");
        checkArgument(getWorkbookOf(from).equals(getWorkbookOf(to)), "the specified Cell is not located in the same Workbook as the target Cell");
        copy(from, to);
        from.getRow().removeCell(from);
        return to;
    }

    /**
     * Returns the value of the specified cell parsed as a {@code Boolean}.
     * <p>
     * This method is {@code null} safe.
     * <p>
     * Note: this method defines a boolean value differently than {@link Boolean#parseBoolean(String) Java}. If the
     * formatted cell value is not equal to the string "true" or "false" (ignoring case and whitespace) this call will
     * result in an exception.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Boolean}
     * @throws IllegalArgumentException if the value of the cell cannot be parsed as a boolean
     */
    public static Boolean parseBoolean(final Cell cell) {
        if (cell == null)
            return null;

        final String value = trim(formatValue(cell));
        if (value.equalsIgnoreCase("true"))
            return true;
        else if (value.equalsIgnoreCase("false"))
            return false;
        else
            throw new IllegalArgumentException();

    }

    /**
     * Returns the value of the specified cell parsed as a {@code Byte}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Byte}
     * @throws NumberFormatException if the value of the cell cannot be parsed as a byte
     */
    public static Byte parseByte(final Cell cell) {
        return cell == null ? null : new Byte(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as a {@code Double}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Double}
     * @throws NumberFormatException if the value of the cell cannot be parsed as a double
     */
    public static Double parseDouble(final Cell cell) {
        return cell == null ? null : new Double(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as a {@code Float}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Float}
     * @throws NumberFormatException if the value of the cell cannot be parsed as a float
     */
    public static Float parseFloat(final Cell cell) {
        return cell == null ? null : new Float(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as an {@code Integer}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as an {@code Integer}
     * @throws NumberFormatException if the value of the cell cannot be parsed as an integer
     */
    public static Integer parseInteger(final Cell cell) {
        return cell == null ? null : new Integer(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as a {@code Long}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Long}
     * @throws NumberFormatException if the value of the cell cannot be parsed as a long
     */
    public static Long parseLong(final Cell cell) {
        return cell == null ? null : new Long(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as a {@code Short}.
     * <p>
     * This method is {@code null} safe.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as a {@code Short}
     * @throws NumberFormatException if the value of the cell cannot be parsed as a short
     */
    public static Short parseShort(final Cell cell) {
        return cell == null ? null : new Short(trim(formatValue(cell)));
    }

    /**
     * Returns the value of the specified cell parsed as an {@code Instant} using the system default time-zone offset.
     * 
     * @param cell the specified cell
     * @return the value of the specified cell parsed as an {@code Instant} using the system default time-zone offset
     */
    public static Instant parseDate(final Cell cell) {
        final Double d = parseDouble(cell);
        if (d == null)
            return null;
        // checkState(d > -Double.MIN_VALUE, "The specified cell cannot be parsed as a Date, %s < -Double.MIN_VALUE", d);
        return DateUtil.getLocalDateTime(d).toInstant(OffsetDateTime.now().getOffset());
    }

    /**
     * Returns the formatted value of the specified cell.
     * <p>
     * This method is {@code null} safe. If the specified cell is {@code null} it will return a {@code null} value.
     * <p>
     * The intention is to retrieve the data in the specified cell the exact way you would see it in Microsoft Excel,
     * regardless of the cell type (e.g. 5.200 would be returned as 5.200 not 5.2).
     * <p>
     * Note: This method is not equivalent to {@link Cell#getStringCellValue()}.
     * 
     * @param cell the specified cell
     * @return the formatted value of the specified cell
     */
    public static String formatValue(final Cell cell) {
        if (cell == null)
            return null;
        // if (cell.getCellType() == CellType.BOOLEAN) return cell.toString().toUpperCase(); // why do we need this?
        return DATA_FORMATTER.formatCellValue(cell);
    }

    /**
     * Returns the row that owns the specified cell. If the cell has been deleted this method will result in an exception.
     * 
     * @param cell the specified cell
     * @return the row that owns the specified cell
     */
    public static Row getRowOf(final Cell cell) {
        checkNotNull(cell, "cell == null");
        return cell.getRow();
    }

    /**
     * Returns the sheet the specified cell belongs to. If the cell has been deleted this method will result in an
     * exception.
     * 
     * @param cell the specified cell
     * @return the sheet the specified cell belongs to
     */
    public static Sheet getSheetOf(final Cell cell) {
        checkNotNull(cell, "cell == null");
        return cell.getSheet();
    }

    /**
     * Return the workbook the specified cell belongs to. If the cell has been deleted this method will result in an
     * exception.
     * 
     * @param cell the specified cell
     * @return the workbook the cell belongs to
     */
    public static Workbook getWorkbookOf(final Cell cell) {
        checkNotNull(cell, "cell == null");
        return Sheets.getWorkbookOf(getSheetOf(cell));
    }

    /**
     * Creates a hyperlink in the specified cell.
     * 
     * @param cell    the specified cell
     * @param type    the type of hyperlink to create
     * @param address the hyperlink address
     * @param label   the label to use for this hyperlink
     * @param value   the text value to be set for the cell
     * @return the specified cell
     */
    public static Cell setHyperlink(final Cell cell, final HyperlinkType type, final String address, final String label, final String value) {
        checkNotNull(cell, "cell == null");
        checkNotNull(address, "address == null");
        checkNotNull(label, "label == null");
        checkNotNull(value, "value == null");
        final Hyperlink link = getWorkbookOf(cell).getCreationHelper().createHyperlink(type);
        link.setAddress(address);
        link.setLabel(label);
        cell.setHyperlink(link);
        cell.setCellValue(value);
        return cell;
    }

    /**
     * Set the style for the cell.
     * <p>
     * Note: the {@code CellStye} object must be created from the workbook where the cell is located.
     * 
     * @param cell  the specified cell
     * @param style the style to set
     * @return the specified cell
     */
    public static Cell setStyle(final Cell cell, final CellStyle style) {
        checkNotNull(cell, "cell == null");
        checkNotNull(style, "style == null");
        cell.setCellStyle(style);
        return cell;
    }

    /**
     * Sets a value for the specified cell.
     * <p>
     * This method is <b>not</b> {@code null} safe. If {@code value} is {@code null} a {@code NullPointerException} will be
     * thrown.
     * <p>
     * If {@code value} is a {@link Number} the cell value will be set to the {@code double} value of the number by first
     * calling {@link Number#doubleValue()} followed by {@link Cell#setCellValue(double)}.
     * <p>
     * If {@code value} is a {@code Boolean}, {@link Calendar}, {@link Date}, {@link LocalDateTime}, or
     * {@link RichTextString} the cell value will be set by calling {@link Cell#setCellValue(boolean)},
     * {@link Cell#setCellValue(Calendar)}, {@link Cell#setCellValue(Date)}, {@link Cell#setCellValue(LocalDateTime)}, or
     * {@link Cell#setCellValue(RichTextString)} respectively.
     * <p>
     * For all other types the cell will be set to {@code value.toString()}. Note if the result of {@code value.toString()}
     * is empty or composed exclusively of whitespace characters according to {@link CharMatcher#WHITESPACE}, it will be
     * ignored (to insert such a string use {@link Cell#setCellValue(String)} directly).
     * 
     * @param cell  the specified cell
     * @param value the value to set
     * @return the specified cell
     * @throws IllegalArgumentException if the value to set is a string which exceeds 32767 characters
     */
    public static Cell setValue(final Cell cell, final Object value) {
        checkNotNull(cell, "cell == null");
        checkNotNull(value, "value == null");
        if (value instanceof Boolean)
            cell.setCellValue((Boolean) value);
        else if (value instanceof Calendar)
            cell.setCellValue((Calendar) value);
        else if (value instanceof Date)
            cell.setCellValue((Date) value);
        else if (value instanceof Number)
            cell.setCellValue(((Number) value).doubleValue());
        else if (value instanceof LocalDateTime)
            cell.setCellValue((LocalDateTime) value);
        else if (value instanceof RichTextString)
            cell.setCellValue((RichTextString) value);
        else {
            final String string = value.toString();
            if (!isWhitespace(string)) {
                if (string.length() > 32767)
                    throw new IllegalArgumentException("value > 32767 characters");
                cell.setCellValue(string);
            }
        }
        return cell;
    }

}