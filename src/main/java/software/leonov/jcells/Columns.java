package software.leonov.jcells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.ss.util.CellReference;

/**
 * Static methods for working with columns of a {@code Sheet}.
 * 
 * @author Zhenya Leonov
 */
public final class Columns {

    private Columns() {
    }

    /**
     * Returns the 0-based index of the specified column.
     * 
     * @param ref the letter reference of the column
     * @return the 0-based index of the specified column
     */
    public static int getIndex(final String ref) {
        checkNotNull(ref, "ref == null");
        return CellReference.convertColStringToIndex(ref);
    }

    /**
     * Returns the letter reference of the column index.
     * 
     * @param index the 0-based index of the column
     * @return the letter reference of the column index
     */
    public static String getReference(final int index) {
        checkArgument(index >= 0, "index < 0");
        return CellReference.convertNumToColString(index);
    }

}
