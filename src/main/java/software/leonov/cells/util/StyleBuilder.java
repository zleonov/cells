package software.leonov.cells.util;

import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * A builder for creating {@link CellStyle}s.
 * <p>
 * Example:
 * <p>
 * 
 * <pre>
 *   import static org.apache.poi.ss.usermodel.IndexedColors.*;
 *   import static org.apache.poi.ss.usermodel.CellStyle.*;
 *
 *   // Creating a new style
 *
 *   final StyleBuilder builder = new StyleBuilder()
 *                         .setWrapText(true)
 *                         .setBorder(BorderStyle.DASH_DOT);
 *
 *   final CellStyle wrappedDashDotStyle = builder.create(workbook);
 *
 *   final CellStyle wrappedBorderThinStyle = builder
 *                        .setBorder(BorderStyle.BORDER_THIN)
 *                        .create(workbook);
 *                        
 *   // Update an existing style
 *   
 *   CellStyle style = workbook.createCellStyle();
 *   ...
 *   new StyleBuilder().setWrapText(false).update(style);
 * </pre>
 * 
 * Builder instances are reusable. It maintains its own state and can create or update multiple {@code CellStyle}
 * instances across different workbooks.
 * <p>
 * <b>Note:</b> A workbook can store a finite number of cell-styles. Be careful not to create identical instances.
 * Styles should be reused whenever possible.
 * 
 * @author Zhenya Leonov
 */
public final class StyleBuilder {

    // Border styles
    private BorderStyle topBorder    = null;
    private BorderStyle bottomBorder = null;
    private BorderStyle leftBorder   = null;
    private BorderStyle rightBorder  = null;

    // Border colors
    private IndexedColors topBorderColor    = null;
    private IndexedColors bottomBorderColor = null;
    private IndexedColors leftBorderColor   = null;
    private IndexedColors rightBorderColor  = null;

    // Data format
    private Short dataFormat = null;

    // Fill colors and pattern
    private IndexedColors   fillBackgroundColor = null;
    private IndexedColors   fillForegroundColor = null;
    private FillPatternType fillPattern         = null;

    // Font
    private Font font = null;

    // Cell alignment
    private HorizontalAlignment horizontalAlignment = null;
    private VerticalAlignment   verticalAlignment   = null;

    // Cell properties
    private Boolean hidden        = null;
    private Short   indention     = null;
    private Boolean locked        = null;
    private Boolean quotePrefixed = null;
    private Short   rotation      = null;
    private Boolean shrinkToFit   = null;
    private Boolean wrapText      = null;

    /**
     * Creates a new {@code StyleBuilder} with no default settings.
     */
    public StyleBuilder() {
    }

    /**
     * Creates a new {@code CellStyle} in the provided workbook based on the current builder settings.
     * 
     * @param workbook the workbook where the cell-style will be created
     * @return a newly-created {@code CellStyle}
     */
    public CellStyle create(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        return update(workbook.createCellStyle());
    }

    /**
     * Creates a new {@code CellStyle} in the provided workbook, initialized with the specified cell-style, and then updated
     * with the current builder settings.
     * 
     * @param workbook  the workbook where the cell-style will be created
     * @param baseStyle the cell-style to use as a base
     * @return the newly-created {@code CellStyle}
     */
    public CellStyle create(final Workbook workbook, final CellStyle baseStyle) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(baseStyle, "baseStyle == null");

        final CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        return update(style);
    }

    /**
     * Updates the provided style with the current builder settings.
     * 
     * @param style the style to update
     * @return the updated style
     */
    public CellStyle update(final CellStyle style) {
        checkNotNull(style, "style == null");
        applyToStyle(style);
        return style;
    }

    /**
     * Clears all settings from this builder, returning it to its initial state.
     * 
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder clear() {
        this.topBorder           = null;
        this.bottomBorder        = null;
        this.leftBorder          = null;
        this.rightBorder         = null;
        this.topBorderColor      = null;
        this.bottomBorderColor   = null;
        this.leftBorderColor     = null;
        this.rightBorderColor    = null;
        this.dataFormat          = null;
        this.fillBackgroundColor = null;
        this.fillForegroundColor = null;
        this.fillPattern         = null;
        this.font                = null;
        this.hidden              = null;
        this.indention           = null;
        this.locked              = null;
        this.quotePrefixed       = null;
        this.rotation            = null;
        this.shrinkToFit         = null;
        this.horizontalAlignment = null;
        this.verticalAlignment   = null;
        this.wrapText            = null;
        return this;
    }

    /**
     * Sets the type of horizontal alignment.
     * 
     * @param align the type of alignment
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setAlignment(final HorizontalAlignment align) {
        checkNotNull(align, "align == null");
        this.horizontalAlignment = align;
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
        this.topBorder    = border;
        this.bottomBorder = border;
        this.leftBorder   = border;
        this.rightBorder  = border;
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

        if (side == BorderSide.TOP)
            this.topBorder = border;
        else if (side == BorderSide.BOTTOM)
            this.bottomBorder = border;
        else if (side == BorderSide.LEFT)
            this.leftBorder = border;
        else
            this.rightBorder = border;

        return this;
    }

    /**
     * Sets the color to use for the border of the entire cell (top/right/bottom/left).
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setBorderColor(final IndexedColors color) {
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
     */
    public StyleBuilder setBorderColor(final BorderSide side, final IndexedColors color) {
        checkNotNull(side, "side == null");
        checkNotNull(color, "color == null");

        if (side == BorderSide.TOP)
            this.topBorderColor = color;
        else if (side == BorderSide.BOTTOM)
            this.bottomBorderColor = color;
        else if (side == BorderSide.LEFT)
            this.leftBorderColor = color;
        else if (side == BorderSide.RIGHT)
            this.rightBorderColor = color;

        return this;
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
        this.dataFormat = fmt;
        return this;
    }

    /**
     * Sets the background fill color.
     * <p>
     * This method works in concert with {@link #setFillForegroundColor(IndexedColors)} and
     * {@link #setFillPattern(FillPatternType)} to produce the desired results.
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see #setFillForegroundColor(IndexedColors)
     */
    public StyleBuilder setFillBackgroundColor(final IndexedColors color) {
        checkNotNull(color, "color == null");
        this.fillBackgroundColor = color;
        return this;
    }

    /**
     * Sets the foreground fill color.
     * <p>
     * This method works in concert with {@link #setFillBackgroundColor(IndexedColors)} and
     * {@link #setFillPattern(FillPatternType)} to produce the desired results.
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     * @see #setFillBackgroundColor(IndexedColors)
     */
    public StyleBuilder setFillForegroundColor(final IndexedColors color) {
        checkNotNull(color, "color == null");
        this.fillForegroundColor = color;
        return this;
    }

    /**
     * Sets the fill pattern of the cell.
     * <p>
     * This method works in concert with {@link #setFillBackgroundColor(IndexedColors)} and
     * {@link #setFillForegroundColor(IndexedColors)} to produce the desired results.
     * 
     * @param fp the fill pattern
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setFillPattern(final FillPatternType fp) {
        checkNotNull(fp, "fp == null");
        this.fillPattern = fp;
        return this;
    }

    /**
     * Sets the {@link #setFillForegroundColor(IndexedColors) foreground} fill color and
     * {@link #setFillPattern(FillPatternType) applies} a {@link FillPatternType#SOLID_FOREGROUND solid foreground} fill
     * pattern.
     * 
     * @param color the color to set
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setSolidFillColor(final IndexedColors color) {
        checkNotNull(color, "color == null");
        this.fillForegroundColor = color;
        this.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return this;
    }

//    /**
//     * Sets the font for this style.
//     * 
//     * @param font the specified font
//     * 
//     * @return this {@code StyleBuilder} instance
//     * @see Workbook#createFont()
//     * @see Workbook#getFontAt(short)
//     * @see CellStyle#getFontIndex()
//     */
//    public StyleBuilder setFont(final Font font) {
//        checkNotNull(font, "font == null");
//        this.font = font;
//        return this;
//    }

    /**
     * Sets the cells using this style to be hidden.
     * 
     * @param hidden specifies whether or not to hide the cells using this style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setHidden(final boolean hidden) {
        this.hidden = hidden;
        return this;
    }

    /**
     * Set the number of spaces to indent the text in the cell.
     * 
     * @param indent number of spaces
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setIndention(final short indent) {
        this.indention = indent;
        return this;
    }

    /**
     * Sets the cells using this style to be locked.
     * 
     * @param locked specifies whether or not to lock the cells using this style
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setLocked(final boolean locked) {
        this.locked = locked;
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
        this.quotePrefixed = treatAsText;
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
        this.rotation = rotation;
        return this;
    }

    /**
     * Sets whether or not the cell should auto-sized to fit its contents.
     * 
     * @param shrinkToFit whether or not the cell should auto-sized to fit its contents
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setShrinkToFit(final boolean shrinkToFit) {
        this.shrinkToFit = shrinkToFit;
        return this;
    }

    /**
     * Sets the type of vertical alignment.
     * 
     * @param align the type of alignment
     * @return this {@code StyleBuilder} instance
     */
    public StyleBuilder setVerticalAlignment(final VerticalAlignment align) {
        checkNotNull(align, "align == null");
        this.verticalAlignment = align;
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
        this.wrapText = wrapped;
        return this;
    }

    private void applyToStyle(final CellStyle style) {
        if (horizontalAlignment != null)
            style.setAlignment(horizontalAlignment);

        if (topBorder != null)
            style.setBorderTop(topBorder);
        if (bottomBorder != null)
            style.setBorderBottom(bottomBorder);
        if (leftBorder != null)
            style.setBorderLeft(leftBorder);
        if (rightBorder != null)
            style.setBorderRight(rightBorder);

        // Apply border colors
        if (topBorderColor != null)
            applyBorderColor(style, BorderSide.TOP, topBorderColor);
        if (bottomBorderColor != null)
            applyBorderColor(style, BorderSide.BOTTOM, bottomBorderColor);
        if (leftBorderColor != null)
            applyBorderColor(style, BorderSide.LEFT, leftBorderColor);
        if (rightBorderColor != null)
            applyBorderColor(style, BorderSide.RIGHT, rightBorderColor);

        if (dataFormat != null)
            style.setDataFormat(dataFormat);

        // Apply fill styles
        if (fillForegroundColor != null)
            applyFillForegroundColor(style, fillForegroundColor);
        if (fillBackgroundColor != null)
            applyFillBackgroundColor(style, fillBackgroundColor);
        if (fillPattern != null)
            style.setFillPattern(fillPattern);

        if (font != null)
            style.setFont(font);
        if (hidden != null)
            style.setHidden(hidden);
        if (indention != null)
            style.setIndention(indention);
        if (locked != null)
            style.setLocked(locked);
        if (quotePrefixed != null)
            style.setQuotePrefixed(quotePrefixed);
        if (rotation != null)
            style.setRotation(rotation);
        if (shrinkToFit != null)
            style.setShrinkToFit(shrinkToFit);
        if (verticalAlignment != null)
            style.setVerticalAlignment(verticalAlignment);
        if (wrapText != null)
            style.setWrapText(wrapText);
    }

    /**
     * Returns a new {@code StyleBuilder} instance populated with the current properties of {@code this}
     * {@code StyleBuilder}.
     * 
     * @return a new {@code StyleBuilder} instance populated with the current properties of {@code this}
     *         {@code StyleBuilder}
     */
    public StyleBuilder newStyleBuilder() {
        final StyleBuilder builder = new StyleBuilder();

        // Border styles
        builder.topBorder    = topBorder;
        builder.bottomBorder = bottomBorder;
        builder.leftBorder   = leftBorder;
        builder.rightBorder  = rightBorder;

        // Border colors
        builder.topBorderColor    = topBorderColor;
        builder.bottomBorderColor = bottomBorderColor;
        builder.leftBorderColor   = leftBorderColor;
        builder.rightBorderColor  = rightBorderColor;

        // Data format
        builder.dataFormat = dataFormat;

        // Fill colors and pattern
        builder.fillBackgroundColor = fillBackgroundColor;
        builder.fillForegroundColor = fillForegroundColor;
        builder.fillPattern         = fillPattern;

        // Font
        builder.font = font;

        // Cell alignment
        builder.horizontalAlignment = horizontalAlignment;
        builder.verticalAlignment   = verticalAlignment;

        // Cell properties
        builder.hidden        = hidden;
        builder.indention     = indention;
        builder.locked        = locked;
        builder.quotePrefixed = quotePrefixed;
        builder.rotation      = rotation;
        builder.shrinkToFit   = shrinkToFit;
        builder.wrapText      = wrapText;

        return builder;
    }

    private void applyBorderColor(final CellStyle style, final BorderSide side, final IndexedColors color) {
        if (color == null)
            return;

        style.setTopBorderColor(color.getIndex());
        style.setBottomBorderColor(color.getIndex());
        style.setLeftBorderColor(color.getIndex());
        style.setRightBorderColor(color.getIndex());
    }

    private void applyFillForegroundColor(final CellStyle style, final IndexedColors color) {
        style.setFillForegroundColor(color.getIndex());
    }

    private void applyFillBackgroundColor(final CellStyle style, final IndexedColors color) {
        style.setFillBackgroundColor(color.getIndex());
    }

}