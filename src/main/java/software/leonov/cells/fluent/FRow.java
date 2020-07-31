package software.leonov.cells.fluent;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.util.Iterator;
import java.util.Optional;
import java.util.concurrent.ExecutionException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.google.common.collect.Iterators;
import com.google.common.collect.Streams;

import software.leonov.common.base.Obj;

/**
 * A representation of a row in a sheet in a Microsoft Excel workbook.
 * 
 * @author Zhenya Leonov
 */
public final class FRow implements Iterable<FCell> {

    private static final Cache<Cell, FCell> cells = CacheBuilder.newBuilder().maximumSize(1000).build();

    private final FSheet fsheet;
    private final Row row;

    FRow(final FSheet fsheet, final Row row) {
        checkNotNull(fsheet, "fsheet == null");
        checkNotNull(row, "row == null");
        this.fsheet = fsheet;
        this.row = row;
    }

    Row delegate() {
        return row;
    }

    final FRow removeCell(final FCell cell) {
        checkNotNull(cell, "cell == null");
        row.removeCell(cell.delegate());
        return this;
    }

    /**
     * Returns the specified cell or {@code null} if the cell is undefined.
     * 
     * @param column the 0-based column number
     * @return the specified cell or {@code null}
     */
    public FCell getCell(final int column) {
        checkArgument(column >= 0, "column index < 0");
        final Cell cell = row.getCell(column);
        try {
            return cell == null ? null : cells.get(cell, () -> new FCell(this, cell));
        } catch (ExecutionException e) {
            throw new AssertionError(e); // cannot happen
        }
    }

    /**
     * Returns the specified cell or {@code null} if the cell is undefined.
     * 
     * @param ref the letter reference of the column
     * @return the specified cell or {@code null}
     */
    public FCell getCell(final String ref) {
        checkNotNull(ref, "ref == null");
        final Cell cell = row.getCell(CellReference.convertColStringToIndex(ref));
        try {
            return cell == null ? null : cells.get(cell, () -> new FCell(this, cell));
        } catch (ExecutionException e) {
            throw new AssertionError(e); // cannot happen
        }
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * <p>
     * The cell-style will be inherited from the default column style. If the default column style is undefined the
     * cell-style will be inherited from the default row style. If neither are defined it will have a {@link CellType#BLANK}
     * style.
     * 
     * @param index the 0-based column index
     * @see Sheet#getColumnStyle(int)
     * @see Row#getRowStyle()
     * @return the specified cell
     */
    public FCell getOrCreateCell(final int index) {
        checkArgument(index >= 0, "index < 0");

        Cell cell = row.getCell(index);

        if (cell == null) {
            cell = row.createCell(index);
            final CellStyle style = Obj.coalesce(row.getSheet().getColumnStyle(index), row.getRowStyle());
            if (style != null)
                cell.setCellStyle(style);
            row.setHeightInPoints(row.getHeightInPoints());
        }

        FCell fcell = cells.getIfPresent(cell);
        if (fcell == null) {
            fcell = new FCell(this, cell);
            cells.put(cell, fcell);
        }

        return fcell;
    }

    /**
     * Returns the specified cell. If the cell does not exist it is created.
     * 
     * @param column the letter reference of the column
     * @return the specified cell
     */
    public FCell getOrCreateCell(final String ref) {
        checkNotNull(ref, "ref == null");
        final int index = CellReference.convertColStringToIndex(ref);
        return getOrCreateCell(index);
    }

    /**
     * Returns the sheet that contains this row.
     * 
     * @return the sheet that contains this row
     */
    public FSheet getSheet() {
        return fsheet;
    }

    /**
     * Returns the workbook that contains this row.
     * 
     * @return the workbook which contains this row
     */
    public FWorkbook getWorkbook() {
        return getSheet().getWorkbook();
    }

    /**
     * Sets the height of this row.
     * 
     * @param height the height to set, in points
     * @return this row
     */
    public FRow setRowHeight(final float height) {
        checkArgument(height > 0, "height < 1");
        row.setHeightInPoints(height);
        return this;
    }

    /**
     * Applies the cell-style to future and existing cells in this row.
     * 
     * @param style the specified cell-style
     * @return this row
     */
    public FRow setStyle(final CellStyle style) {
        checkNotNull(style, "style == null");
        row.forEach(cell -> cell.setCellStyle(style));
        row.setRowStyle(style);
        return this;
    }

//    public Iterable<FCell> skipBlankCells(final Row row) {
//        checkNotNull(row, "row == null");
//        return Iterables.filter(this, cell -> !isWhitespace(cell.formatValue()));
//    }

    /**
     * Returns the 1-based index of the last cell in this row or an empty {@code Optional} if the row has no
     * defined cells.
     * 
     * @return the 1-based index of the last cell in this row or an empty {@code Optional} if the row has no
     *         defined cells
     */
    public Optional<Integer> getLastCellIndex() {
        final int i = row.getLastCellNum();
        return i == -1 ? Optional.empty() : Optional.of(i);
    }

    @Override
    public Iterator<FCell> iterator() {
        return Iterators.transform(row.iterator(), cell -> new FCell(this, cell));
    }
    
}