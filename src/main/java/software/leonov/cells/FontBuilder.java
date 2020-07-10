package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * A builder of {@link Font}s.
 * <p>
 * Example:
 * <p>
 * 
 * <pre>
 *   import static org.apache.poi.ss.usermodel.IndexedColors.*;
 *   import static org.apache.poi.ss.usermodel.Font.*;
 *   
 *   final FontBuilder builder = new FontBuilder(workbook).setUnderline(DOUBLE).setItalic(true);
 *   
 *   final Font underlinedItalic     = builder.build();
 *   final Font underlinedItalicBold = builder.setBold(true).build();
 *   final Font strikeoutItalic      = builder.setStrikeout(true).setBold(false).build();
 * </pre>
 * 
 * Builder instances are reusable. It is safe to call {@link #build()} multiple times to obtain multiple {@code Font}
 * instances.
 * <p>
 * <b>Note:</b> A workbook can store a finite number of fonts. Be careful not to create identical instances. Fonts
 * should be reused whenever possible.
 * 
 * @author Zhenya Leonov
 */
public final class FontBuilder {

    private final Font font;
    private final Workbook workbook;

    /**
     * Creates a new {@code FontBuilder} using the specified workbook.
     * 
     * @param workbook the workbook where the {@code Font} returned by this builder will reside
     */
    public FontBuilder(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        this.workbook = workbook;
        this.font = workbook.createFont();
    }

    /**
     * Creates a new {@code FontBuilder} initialized with the specified font.
     * 
     * @param workbook the workbook where the font returned by this builder will reside
     * @param font     the base font to start with
     */
    public FontBuilder(final Workbook workbook, final Font font) {
        checkNotNull(font, "font == null");
        checkNotNull(workbook, "workbook == null");
        this.workbook = workbook;
        this.font = workbook.createFont();
        copy(font, this.font);
    }

    /**
     * Returns a newly-created {@code Font} based on the contents of this builder.
     * 
     * @return a newly-created {@code Font} based on the contents of this builder
     */
    public Font build() {
        Font newFont = workbook.createFont();
        copy(font, newFont);
        return font;
    }

    /**
     * Sets whether or not this font is in bold.
     * 
     * @param bold whether or not this font is in bold
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setBold(final boolean bold) {
        font.setBold(bold);
        return this;
    }

    /**
     * Sets character-set to use.
     * 
     * @param charset the character-set to use
     * @return this {@code FontBuilder} instance
     * @see Font#ANSI_CHARSET
     * @see Font#DEFAULT_CHARSET
     * @see Font#SYMBOL_CHARSET
     */
    public FontBuilder setCharSet(final int charset) { // what about XSSFFont.setCharSet(FontCharset)
        font.setCharSet(charset);
        return this;
    }

    /**
     * Sets the color for the font.
     * 
     * @param color the color to set
     * @return this {@code FontBuilder} instance
     * @see HSSFColor
     * @see XSSFColor
     */
    public FontBuilder setColor(final Color color) {
        checkNotNull(color, "color == null");

        if (color instanceof XSSFColor) {
            checkArgument(font instanceof XSSFFont, "XSSFColor is not compatible with %s", font.getClass().getSimpleName());
            ((XSSFFont) font).setColor((XSSFColor) color);
        } else
            font.setColor(Colors.getIndex(color));
        return this;
    }

    /**
     * Sets the font height in units of 1/20th of a point.
     * 
     * @param height height in 1/20ths of a point
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontHeight(final short height) {
        font.setFontHeight(height);
        return this;
    }

    /**
     * Sets the font height.
     * 
     * @param height the font height in the familiar unit of measure - points (10, 12, 14)
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontHeightInPoints(final short height) {
        font.setFontHeightInPoints(height);
        return this;
    }

    /**
     * Set the name of the font (e.g. Arial, Times New Roman). Use {@link CommonFont#getFontName()} for common cross
     * platform fonts.
     * 
     * @param font the name of the font to use
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontName(final String font) {
        checkNotNull(font, "font == null");
        this.font.setFontName(font);
        return this;
    }

    /**
     * Sets whether or not to make the font italic.
     * 
     * @param italic to italicize or not
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setItalic(final boolean italic) {
        font.setItalic(italic);
        return this;
    }

    /**
     * Sets whether or not to use a strikeout horizontal line.
     * 
     * @param strikeout to strikeout or not
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setStrikeout(final boolean strikeout) {
        font.setStrikeout(strikeout);
        return this;
    }

    /**
     * Sets normal, super, or subscript.
     * 
     * @param offset the type use (none, super, sub)
     * @return this {@code FontBuilder} instance
     * @see Font#SS_NONE
     * @see Font#SS_SUPER
     * @see Font#SS_SUB
     */
    public FontBuilder setTypeOffset(final TypeOffset offset) {
        font.setTypeOffset(offset.getShortValue());
        return this;
    }

    /**
     * Sets type of text underlining to use.
     * 
     * @param underline the type of underline
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setUnderline(final FontUnderline underline) {
        checkNotNull(underline, "underline == null");
        font.setUnderline(underline.getByteValue());
        return this;
    }

    private void copy(final Font from, final Font to) {
        to.setBold(from.getBold());
        to.setCharSet(from.getCharSet());
        to.setColor(from.getColor());
        to.setFontHeight(from.getFontHeight());
        to.setFontHeightInPoints(from.getFontHeightInPoints());
        to.setFontName(from.getFontName());
        to.setItalic(from.getItalic());
        to.setStrikeout(from.getStrikeout());
        to.setTypeOffset(from.getTypeOffset());
        to.setUnderline(from.getUnderline());
    }

}
