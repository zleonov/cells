package software.leonov.cells.fluent;

import static com.google.common.base.Preconditions.checkNotNull;
import static software.leonov.common.base.Str.isWhitespace;
import static software.leonov.common.base.Str.trim;
import static software.leonov.common.base.Str.whitespaceToNull;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * A representation of a cell in a row in a sheet in a Microsoft Excel workbook.
 * 
 * @author Zhenya Leonov
 */
public final class FCell {

    /**
     * The total number of characters that a cell can contain
     */
    public static final int MAX_CELL_SIZE = 32767;

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    private final FRow row;
    private final Cell cell;

    FCell(final FRow row, final Cell cell) {
        checkNotNull(row, "row == null");
        checkNotNull(cell, "row == null");
        this.row = row;
        this.cell = cell;
    }

    Cell delegate() {
        return cell;
    }

    public FCell setActive() {
        cell.setAsActiveCell();
        return this;
    }

    /**
     * Returns the formatted value of this cell.
     * <p>
     * The intention is to retrieve the data in the specified cell the exact way you would see it in Microsoft Excel,
     * regardless of the cell type (e.g. 5.200 would be returned as 5.200 not 5.2).
     * <p>
     * Note: This method is not equivalent to {@link Cell#getStringCellValue()}.
     * 
     * @return the formatted value of this cell or {@code null} if the cell is empty, has no value, or contains only
     *         whitespace characters
     */
    public String formatValue() {
        // if (cell.getCellType() == CellType.BOOLEAN) return cell.toString().toUpperCase(); // why do we need this?
        return whitespaceToNull(DATA_FORMATTER.formatCellValue(cell));
    }

    /**
     * Returns the row this cell belongs to.
     * 
     * @return the row this cell belongs to
     */
    public FRow getRow() {
        return row;
    }

    /**
     * Returns the type of this cell.
     * 
     * @return the type of this cell
     */
    public CellType getType() {
        return cell.getCellType();
    }

    public FCell setValue(final Object value) {
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
        return this;
    }

    // style

    public FCell setStyle(final CellStyle style) {
        checkNotNull(style, "style == null");
        cell.setCellStyle(style);
        return this;
    }

    public CellStyle getStyle() {
        return cell.getCellStyle();
    }

    // comment

    public FCell setComment(final Comment comment) {
        checkNotNull(comment, "comment == null");
        cell.setCellComment(comment);
        return this;
    }

    public Comment getComment() {
        return cell.getCellComment();
    }

    public FCell removeComment() {
        cell.removeCellComment();
        return this;
    }

    // hyperlink

    public FCell setHyperlink(final HyperlinkType type, final String address, final String label, final String value) {
        checkNotNull(type,    "type == null");
        checkNotNull(address, "address == null");
        checkNotNull(label,   "label == null");
        checkNotNull(value,   "value == null");
        final Hyperlink link = getRow().getSheet().getWorkbook().delegate().getCreationHelper().createHyperlink(type);
        link.setAddress(address);
        link.setLabel(label);
        cell.setHyperlink(link);
        cell.setCellValue(value);
        return this;
    }

    public Hyperlink getHyperlink() {
        return cell.getHyperlink();
    }

    public FCell removeHyperlink() {
        cell.removeHyperlink();
        return this;
    }

    // parse methods

    /**
     * Returns the value of this cell parsed as a boolean.
     * <p>
     * Note: this method defines a boolean value differently than {@link Boolean#parseBoolean(String) Java}. If the
     * formatted cell value is not equal to the string "true" or "false" (ignoring case and whitespace) this call will
     * result in an exception.
     * 
     * @return the value of this cell parsed as a {@code Boolean}
     * @throws IllegalArgumentException if the value of the cell cannot be parsed as a boolean
     */
    public boolean parseBoolean() {
        final String value = trim(formatValue());
        if (value.equalsIgnoreCase("true"))
            return true;
        else if (value.equalsIgnoreCase("false"))
            return false;
        else
            throw new IllegalArgumentException();
    }

    /**
     * Returns the value of this cell parsed as a byte.
     * 
     * @return the value of this cell parsed as a byte
     * @throws NumberFormatException if the value of the cell cannot be parsed as an byte
     */
    public byte parseByte() {
        return Byte.parseByte(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as a double.
     * 
     * @return the value of this cell parsed as a double
     * @throws NumberFormatException if the value of the cell cannot be parsed as a double
     */
    public double parseDouble() {
        return Double.parseDouble(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as a float
     * 
     * @return the value of this cell parsed as a float
     * @throws NumberFormatException if the value of the cell cannot be parsed as a float
     */
    public float parseFloat() {
        return Float.parseFloat(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as an int.
     * 
     * @return the value of this cell parsed as an int
     * @throws NumberFormatException if the value of the cell cannot be parsed as an int
     */
    public int parseInt() {
        return Integer.parseInt(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as a long.
     * 
     * @return the value of this cell parsed as a long
     * @throws NumberFormatException if the value of the cell cannot be parsed as a long
     */
    public long parseLong() {
        return Long.parseLong(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as a short.
     * 
     * @return the value of this cell parsed as a short
     * @throws NumberFormatException if the value of the cell cannot be parsed as a short
     */
    public short parseShort() {
        return Short.parseShort(trim(formatValue()));
    }

    /**
     * Returns the value of this cell parsed as an {@code Instant} using the system default time-zone offset.
     * 
     * @return the value of this cell parsed as an {@code Instant} using the system default time-zone offset
     */
    public Instant parseDate() {
        final Double d = parseDouble();
        // checkState(d > -Double.MIN_VALUE, "The specified cell cannot be parsed as a Date, %s < -Double.MIN_VALUE", d);
        return DateUtil.getLocalDateTime(d).toInstant(OffsetDateTime.now().getOffset());
    }

}