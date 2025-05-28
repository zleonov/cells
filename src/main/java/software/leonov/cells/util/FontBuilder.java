package software.leonov.cells.util;

import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * A builder of {@link Font}s.
 * <p>
 * Example:
 * 
 * <pre>
 *   import static org.apache.poi.ss.usermodel.IndexedColors.*;
 *   import static org.apache.poi.ss.usermodel.Font.*;
 *   
 *   // Create a new fonts
 *   
 *   final FontBuilder builder = new FontBuilder();
 *   
 *   final Font underlinedItalic     = builder.setUnderline(DOUBLE).setItalic(true).create(workbook);
 *   final Font underlinedItalicBold = builder.setBold(true).create(workbook);
 *   
 *   // Update an existing font
 *   
 *   Font font = workbook.createFont();
 *   ...   
 *   new FontBuilder().setStrikeout(true).setBold(false).update(font);
 * </pre>
 * 
 * Builder instances are reusable. It maintains its own state and can create or update multiple {@code Font} instances
 * across different workbooks.
 * <p>
 * <b>Note:</b> A workbook can store a finite number of fonts. Be careful not to create identical instances. Fonts
 * should be reused whenever possible.
 * 
 * @author Zhenya Leonov
 */
public final class FontBuilder {

    private Boolean bold    = null;
    private Integer charset = null;

    private IndexedColors color = null;

    private Short   fontHeight         = null;
    private Short   fontHeightInPoints = null;
    private String  fontName           = null;
    private Boolean italic             = null;
    private Boolean strikeout          = null;
    private Short   typeOffset         = null;
    private Byte    underline          = null;

    /**
     * Creates a new {@code FontBuilder} with no default settings.
     */
    public FontBuilder() {
    }

    /**
     * Creates a new {@code Font} in the provided workbook based on the current settings.
     * 
     * @param workbook the workbook where the font will be created
     * @return a newly-created {@code Font}
     */
    public Font create(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        final Font font = workbook.createFont();
        applyToFont(font);
        return font;
    }

    /**
     * Creates a new {@code Font} in the provided workbook, initialized with the properties of the provided font, and then
     * updated with the current builder settings.
     * 
     * @param workbook the workbook where the font will be created
     * @param baseFont the font to use as a base
     * @return the newly-created {@code Font}
     */
    public Font create(final Workbook workbook, final Font baseFont) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(baseFont, "baseFont == null");

        final Font font = workbook.createFont();

        font.setBold(baseFont.getBold());
        font.setCharSet(baseFont.getCharSet());
        font.setColor(baseFont.getColor());
        font.setFontHeight(baseFont.getFontHeight());
        font.setFontHeightInPoints(baseFont.getFontHeightInPoints());
        font.setFontName(baseFont.getFontName());
        font.setItalic(baseFont.getItalic());
        font.setStrikeout(baseFont.getStrikeout());
        font.setTypeOffset(baseFont.getTypeOffset());
        font.setUnderline(baseFont.getUnderline());

        applyToFont(font);
        return font;
    }

    /**
     * Updates the provided font with the current builder settings.
     * 
     * @param font the font to update
     */
    public void update(final Font font) {
        checkNotNull(font, "font == null");
        applyToFont(font);
    }

    /**
     * Sets whether or not this font is in bold.
     * 
     * @param bold whether or not this font is in bold
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setBold(final boolean bold) {
        this.bold = bold;
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
    public FontBuilder setCharSet(final int charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Sets the color for the font.
     * 
     * @param color the color to set
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setColor(final IndexedColors color) {
        checkNotNull(color, "color == null");
        this.color = color;
        return this;
    }

    /**
     * Sets the font height in units of 1/20th of a point.
     * 
     * @param height height in 1/20ths of a point
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontHeight(final short height) {
        this.fontHeight = height;
        return this;
    }

    /**
     * Sets the font height.
     * 
     * @param height the font height in the familiar unit of measure - points (10, 12, 14)
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontHeightInPoints(final short height) {
        this.fontHeightInPoints = height;
        return this;
    }

    /**
     * Set the name of the font (e.g. Arial, Times New Roman). Use {@link CommonFont#getFontName()} for common cross
     * platform fonts.
     * 
     * @param fontName the name of the font to use
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setFontName(final String fontName) {
        checkNotNull(fontName, "fontName == null");
        this.fontName = fontName;
        return this;
    }

    /**
     * Sets whether or not to make the font italic.
     * 
     * @param italic to italicize or not
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setItalic(final boolean italic) {
        this.italic = italic;
        return this;
    }

    /**
     * Sets whether or not to use a strikeout horizontal line.
     * 
     * @param strikeout to strikeout or not
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder setStrikeout(final boolean strikeout) {
        this.strikeout = strikeout;
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
        checkNotNull(offset, "offset == null");
        this.typeOffset = offset.getShortValue();
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
        this.underline = underline.getByteValue();
        return this;
    }

    /**
     * Clears all settings from this builder, returning it to its initial state.
     * 
     * @return this {@code FontBuilder} instance
     */
    public FontBuilder clear() {
        this.bold               = null;
        this.charset            = null;
        this.color              = null;
        this.fontHeight         = null;
        this.fontHeightInPoints = null;
        this.fontName           = null;
        this.italic             = null;
        this.strikeout          = null;
        this.typeOffset         = null;
        this.underline          = null;
        return this;
    }

    /**
     * Returns a new {@code FontBuilder} instance populated with the current properties of {@code this} {@code FontBuilder}.
     * 
     * @return a new {@code FontBuilder} instance populated with the current properties of {@code this} {@code FontBuilder}
     */
    public FontBuilder newFontBuilder() {
        final FontBuilder builder = new FontBuilder();

        builder.bold               = bold;
        builder.charset            = charset;
        builder.color              = color;
        builder.fontHeight         = fontHeight;
        builder.fontHeightInPoints = fontHeightInPoints;
        builder.fontName           = fontName;
        builder.italic             = italic;
        builder.strikeout          = strikeout;
        builder.typeOffset         = typeOffset;
        builder.underline          = underline;

        return builder;
    }

    private void applyToFont(final Font font) {
        if (bold != null)
            font.setBold(bold);
        if (charset != null)
            font.setCharSet(charset);
        if (color != null)
            setColor(font, color);
        if (fontHeight != null)
            font.setFontHeight(fontHeight);
        if (fontHeightInPoints != null)
            font.setFontHeightInPoints(fontHeightInPoints);
        if (fontName != null)
            font.setFontName(fontName);
        if (italic != null)
            font.setItalic(italic);
        if (strikeout != null)
            font.setStrikeout(strikeout);
        if (typeOffset != null)
            font.setTypeOffset(typeOffset);
        if (underline != null)
            font.setUnderline(underline);
    }

    private static void setColor(final Font font, final IndexedColors color) {
        font.setColor(color.getIndex());
    }

}