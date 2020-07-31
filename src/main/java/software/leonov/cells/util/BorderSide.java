package software.leonov.cells.util;

import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

/**
 * The sides of a cell border. This enum is a copy of {@link XSSFCellBorder.BorderSide}.
 * 
 * @author Zhenya Leonov
 */
public enum BorderSide {
    /**
     * Top side of the cell.
     */
    TOP,

    /**
     * Right side of the cell.
     */
    RIGHT,

    /**
     * Bottom side of the cell.
     */
    BOTTOM,

    /**
     * Left side of the cell.
     */
    LEFT
}
