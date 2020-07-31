package software.leonov.cells.util;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.google.common.base.Preconditions;

import software.leonov.cells.fluent.FWorkbook;

/**
 * A builder for creating {@link CellStyle}s.
 * <p>
 * This builder employees the method chaining paradigm.
 * <p>
 * Example:
 * <p>
 * 
 * <pre>
 *   import static org.apache.poi.ss.usermodel.IndexedColors.*;
 *   import static org.apache.poi.ss.usermodel.CellStyle.*;
 *
 *   final StyleBuilder builder = new StyleBuilder(workbook)
 *                         .setWrapText(true)
 *                         .setBorder(BorderStyle.DASH_DOT);
 *
 *   final CellStyle wrappedDashDotStyle = builder.build();
 *
 *   final CellStyle wrappedBorderThinStyle = builder
 *                        .setBorder(BorderStyle.BORDER_THIN)
 *                        .build();
 * </pre>
 * 
 * Builder instances are reusable. It is safe to call {@link #build()} multiple times to obtain multiple
 * {@code CellStyle} instances.
 * <p>
 * <b>Note:</b> A workbook can store a finite number of cell-styles. Be careful not to create identical instances.
 * Styles should be reused whenever possible.
 * 
 * @author Zhenya Leonov
 */
public final class StyleBuilder {

    private final CellStyle style;
    private final Workbook workbook;

    /**
     * Creates a new {@code StyleBuilder} using the specified workbook.
     * <p>
     * Note: This builder does not create a new {@code Font} object for the underlying cell-style. Modifying the default
     * font by any other means than using a {@link FontBuilder} may affect multiple cell-styles.
     * 
     * @param workbook the workbook where the cell-style returned by this builder will reside
     */
    public StyleBuilder(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        this.style = workbook.createCellStyle();
        this.workbook = workbook;
    }
    
    public StyleBuilder(final FWorkbook workbook) {
        checkNotNull(workbook, "workbook == null");
        this.style = workbook.delegate().createCellStyle();
        this.workbook = workbook.delegate();
    }

    /**
     * Creates a new {@code StyleBuilder} initialized with the specified cell-style. The given {@code CellStyle} object will
     * be copied and left unmodified.
     * <p>
     * Note: The {@code Font} object specified by the underlying cell-style itself is not duplicated. Modifying the font by
     * any other means than using a {@link FontBuilder} may affect multiple cell-styles.
     * 
     * @param workbook the workbook where the cell-style returned by this builder will reside
     * @param style    the base cell-style to start with
     */
    public StyleBuilder(final Workbook workbook, final CellStyle style) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(style, "style == null");
        this.workbook = workbook;
        this.style = workbook.createCellStyle();
        this.style.cloneStyleFrom(style);
    }
    
    public StyleBuilder(final FWorkbook workbook, final CellStyle style) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(style, "style == null");
        this.workbook = workbook.delegate();
        this.style = workbook.delegate().createCellStyle();
        this.style.cloneStyleFrom(style);
    }

    /**
     * Returns a newly-created {@code CellStyle} based on the contents of this builder.
     * 
     * @return a newly-created {@code CellStyle} based on the contents of this builder
     */
    public CellStyle build() {
        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        return newStyle;
    }

    /**
     * Returns a font builder initialized with the font used by the underlying cell-style.
     * 
     * @return a font builder initialized with the font used by the underlying cell-style
     */
    public FontBuilder getFontBuilder() {
        return new FontBuilder(workbook, workbook.getFontAt(style.getFontIndexAsInt()));
    }

    /**
     * Sets the type of horizontal alignment.
     * 
     * @param align the type of alignment
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setAlignment(final HorizontalAlignment align) {
        checkNotNull(align, "alight == null");
        style.setAlignment(align);
        return this;
    }

    /**
     * Sets the type of border to use for the entire cell (top/bottom/right/left).
     * 
     * @param border the border style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setBorder(final BorderStyle border) {
        checkNotNull(border, "border == null");
        for (final BorderSide side : BorderSide.values())
            setBorder(side, border);
        return this;
    }

    /**
     * Sets the type of border to use for the specified {@code BorderSide} of the cell.
     * 
     * @param side   which side of the cell to set
     * @param border the border style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setBorder(final BorderSide side, final BorderStyle border) {
        checkNotNull(side, "side == null");
        checkNotNull(border, "border == null");
        switch (side) {
        case TOP:
            style.setBorderTop(border);
            break;
        case BOTTOM:
            style.setBorderBottom(border);
            break;
        case LEFT:
            style.setBorderLeft(border);
            break;
        case RIGHT:
            style.setBorderRight(border);
            break;
        }
        return this;
    }

    /**
     * Sets the color to use for the border of the entire cell (top/right/bottom/left).
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see Colors
     * @see HSSFColor
     * @see XSSFColor
     */
    public StyleBuilder setBorderColor(final Color color) {
        checkNotNull(color, "color == null");

        for (final BorderSide side : BorderSide.values())
            setBorderColor(side, color);

        return this;
    }

    /**
     * Sets the color to use for the specified {@code BorderSide} of the cell.
     * 
     * @param side  which side of the cell to set
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see Colors
     * @see HSSFColor
     * @see XSSFColor
     */
    public StyleBuilder setBorderColor(final BorderSide side, final Color color) {
        checkNotNull(side, "side == null");
        checkNotNull(color, "color == null");

        if (color instanceof XSSFColor) {
            checkArgument(style instanceof XSSFCellStyle, "XSSFColor is not compatible with %s", style.getClass().getSimpleName());
            return setBorderColor(side, (XSSFColor) color);
        } else
            return setBorderColor(side, Colors.getIndex(color));
    }

    /**
     * Sets the data format (must be a valid format).
     * 
     * @param fmt the data format to set
     * 
     * @return this {@code StyleBuilder} instance
     * @see DataFormat
     * @see CellStyle#getDataFormat()
     */
    public StyleBuilder setDataFormat(final short fmt) {
        style.setDataFormat(fmt);
        return this;
    }

    /**
     * Sets the background fill color.
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see #setFillForegroundColor(Color)
     * @see Colors
     * @see HSSFColor
     * @see XSSFColor
     */
    public StyleBuilder setFillBackgroundColor(final Color color) {
        checkNotNull(color, "color == null");
        if (color instanceof XSSFColor) {
            checkArgument(style instanceof XSSFCellStyle, "XSSFColor is not compatible with %s", style.getClass().getSimpleName());
            ((XSSFCellStyle) style).setFillBackgroundColor((XSSFColor) color);
        } else
            style.setFillBackgroundColor(Colors.getIndex(color));
        return this;
    }

    /**
     * Sets the foreground fill color. The foreground fill color must be set prior to
     * {@link #setFillBackgroundColor(Color)}.
     * <p>
     * This method works in concert with {@link #setFillPattern(FillPatternType)} to produce the desired results.
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see #setFillBackgroundColor(Color)
     * @see Colors
     * @see HSSFColor
     * @see XSSFColor
     */
    public StyleBuilder setFillForegroundColor(final Color color) {
        checkNotNull(color, "color == null");
        if (color instanceof XSSFColor) {
            checkArgument(style instanceof XSSFCellStyle, "XSSFColor is not compatible with %s", style.getClass().getSimpleName());
            ((XSSFCellStyle) style).setFillForegroundColor((XSSFColor) color);
        } else
            style.setFillForegroundColor(Colors.getIndex(color));
        return this;
    }

    /**
     * Sets the fill pattern of the cell.
     * 
     * @param fp the fill pattern
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setFillPattern(final FillPatternType fp) {
        checkNotNull(fp, "fp == null");
        style.setFillPattern(fp);
        return this;
    }

    /**
     * Sets the font for this style.
     * 
     * @param font the specified font
     * 
     * @return this {@code StyleBuilder} instance
     * @see Workbook#createFont()
     * @see Workbook#getFontAt(short)
     * @see CellStyle#getFontIndex()
     */
    public StyleBuilder setFont(final Font font) {
        style.setFont(font);
        return this;
    }

    /**
     * Sets the cells using this style to be hidden.
     * 
     * @param hidden specifies whether or not to hide the cells using this style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setHidden(final boolean hidden) {
        style.setHidden(hidden);
        return this;
    }

    /**
     * Set the number of spaces to indent the text in the cell.
     * 
     * @param indent number of spaces
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setIndention(final short indent) {
        style.setIndention(indent);
        return this;
    }

    /**
     * Sets the cells using this style to be locked.
     * 
     * @param locked specifies whether or not to lock the cells using this style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setLocked(final boolean locked) {
        style.setLocked(locked);
        return this;
    }

    /**
     * Sets whether or not Microsoft Excel should treat the cell value as text even if it can be parsed as a number or a
     * formula.
     * 
     * @param treatAsText Sets whether or not Microsoft Excel should treat the cell value as text even if it can be parsed
     *                    as a number or a formula
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setQuotePrefixed(final boolean treatAsText) {
        style.setQuotePrefixed(treatAsText);
        return this;
    }

    /**
     * Sets the degree of rotation for the text in the cell.
     * <p>
     * <b>Note:</b> Rotation ranges can be set in values of -90 and 90 degrees or 0 and 180 degrees, depending on the
     * underlying cell-style appropriate conversion will be performed.
     * 
     * @param rotation degrees (between -90 and 90 degrees) or (between 0 and 180 degrees)
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setRotation(final short rotation) {
        style.setRotation(rotation);
        return this;
    }

    /**
     * Sets whether or not the cell should auto-sized to fit its contents.
     * 
     * @param shrinkToFit whether or not the cell should auto-sized to fit its contents
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setShrinkToFit(final boolean shrinkToFit) {
        style.setShrinkToFit(shrinkToFit);
        return this;
    }

    /**
     * Sets the type of vertical alignment.
     * 
     * @param align the type of alignment
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setVerticalAlignment(final VerticalAlignment align) {
        checkNotNull(align, "alight == null");
        style.setVerticalAlignment(align);
        return this;
    }

    /**
     * Sets whether the text should be wrapped. Setting this flag to true make all content visible within a cell by
     * displaying it on multiple lines.
     * 
     * @param wrapped specifies whether or not to wrap the text
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setWrapText(final boolean wrapped) {
        style.setWrapText(wrapped);
        return this;
    }

    private StyleBuilder setBorderColor(final BorderSide side, final short index) {
        switch (side) {
        case TOP:
            style.setTopBorderColor(index);
            break;
        case BOTTOM:
            style.setBottomBorderColor(index);
            break;
        case LEFT:
            style.setLeftBorderColor(index);
            break;
        case RIGHT:
            style.setRightBorderColor(index);
            break;
        }
        return this;
    }

    private StyleBuilder setBorderColor(final BorderSide side, final XSSFColor color) {
        switch (side) {
        case TOP:
            ((XSSFCellStyle) style).setTopBorderColor(color);
            break;
        case BOTTOM:
            ((XSSFCellStyle) style).setBottomBorderColor(color);
            break;
        case LEFT:
            ((XSSFCellStyle) style).setLeftBorderColor(color);
            break;
        case RIGHT:
            ((XSSFCellStyle) style).setRightBorderColor(color);
            break;
        }
        return this;
    }

}
