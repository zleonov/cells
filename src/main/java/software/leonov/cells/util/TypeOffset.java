package software.leonov.cells.util;

import org.apache.poi.ss.usermodel.Font;

/**
 * TypeOffset formatting options.
 * 
 * @author Zheny Leonov
 */
public enum TypeOffset {

    /**
     * Superscript type.
     */
    SUPERSCRPIT(Font.SS_SUPER),

    /**
     * Subscript type.
     */
    SUBSCRIPT(Font.SS_SUB),

    /**
     * Normal type.
     */
    NONE(Font.SS_NONE);

    private final short offset;

    private TypeOffset(final short offset) {
        this.offset = offset;
    }

    /**
     * Returns the value of this {@code TypeOffset} compatible with {@link Font#setTypeOffset(short)}.
     * 
     * @return the value of this {@code TypeOffset} compatible with {@link Font#setTypeOffset(short)}
     */
    public short getShortValue() {
        return offset;
    }

}
