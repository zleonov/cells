package software.leonov.cells.fluent;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.concurrent.ExecutionException;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.IntStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.google.common.collect.Iterators;
import com.google.common.io.MoreFiles;

import software.leonov.cells.Workbooks;

/**
 * High level representation of a Microsoft Excel workbook.
 * 
 * @author Zhenya Leonov
 */
public final class FWorkbook implements Iterable<FSheet> {

    private static final Logger logger = Logger.getLogger(FWorkbook.class.getName());

    private static final Cache<Sheet, FSheet> sheets = CacheBuilder.newBuilder().maximumSize(1000).build();

    private Workbook workbook;

    private FWorkbook(final Workbook workbook) {
        this.workbook = workbook;
    }

    public Workbook delegate() {
        return workbook;
    }

    /**
     * Specifies which Microsoft Excel format to use.
     */
    public static enum Format {
        /**
         * The <i>xls</i> <a target="_blank" href= "http://en.wikipedia.org/wiki/Microsoft_Excel#Binary">Excel Binary File
         * Format</a> supported since Microsoft Office 2003.
         */
        BINARY,

        /**
         * The <i>xlsx</i> <a target="_blank" href= "https://en.wikipedia.org/wiki/Office_Open_XML">Office Open XML</a> format
         * supported starting with Microsoft Office 2007.
         */
        OFFICE_OPEN_XML,

        /**
         * <b>Using this format is error prone and not recommended for developers who are making casual use of the API.</b>
         * <p>
         * The streaming version of {@link #OFFICE_OPEN_XML Office Open XML}. See {@link SXSSFWorkbook} and
         * {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new SXSSFWorkbook(XSSFWorkbook)} for more information.
         * <p>
         * <b>Note:</b> Allocated temporary files must be cleaned up explicitly by calling the {@link #dispose(Workbook)}.
         * <p>
         * See the {@link Workbooks#newWorkbook(Format) newWorkbook(Format)} and {@link Workbooks#open(Path, Format) open(Path,
         * Format)} methods for further details.
         */
        STREAMING_OFFICE_OPEN_XML;

    }

    /**
     * Clones the specified sheet.
     * 
     * @param name the name of the target sheet
     * @return the target sheet
     */
    public static FSheet cloneSheet(final FSheet sheet, final String name) {
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final Workbook workbook = sheet.delegate().getWorkbook();
        final Sheet target = workbook.cloneSheet(sheet.getWorkbook().delegate().getSheetIndex(sheet.delegate()));
        final FSheet clone = new FSheet(sheet.getWorkbook(), target);
        return clone.setSheetName(name);
    }

    /**
     * Creates a new workbook with an empty sheet.
     * <p>
     * Workbooks can be created in the classic {@link Format#BINARY Excel Binary File Format} {@code xls} format or the
     * {@link Format#OFFICE_OPEN_XML Office Open XML} {@code xlsx} format.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information.
     * 
     * @param format specifies which workbook format to use
     * @return a new workbook with an empty sheet
     */
    public static FWorkbook newWorkbook(final Format format) {
        checkNotNull(format, "format == null");
        return newWorkbook(format, 1);
    }

    /**
     * Creates a new workbook and adds the specified number of empty sheets.
     * <p>
     * Workbooks can be created in the classic {@link Format#BINARY Excel Binary File Format} {@code xls} format or the
     * {@link Format#OFFICE_OPEN_XML Office Open XML} {@code xlsx} format.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information.
     * 
     * @param format  specifies which workbook format to use
     * @param nsheets the number of sheets to create in the workbook
     * @return a new workbook containing the specified number of empty sheets named <i>Sheet1</i>, <i>Sheet2</i>,
     *         <i>Sheet3</i>, etc...
     */
    public static FWorkbook newWorkbook(final Format format, final int nsheets) {
        checkNotNull(format, "format == null");
        checkArgument(nsheets >= 0, "nsheets < 0");
        final Workbook workbook = format == Format.BINARY ? new HSSFWorkbook() : format == Format.OFFICE_OPEN_XML ? new XSSFWorkbook() : new SXSSFWorkbook();
        IntStream.range(1, nsheets + 1).forEach(i -> workbook.createSheet("Sheet" + i));

        if (nsheets > 0)
            workbook.setActiveSheet(0);

        return new FWorkbook(workbook);
    }

    public boolean dispose() {
        if (workbook instanceof SXSSFWorkbook)
            return ((SXSSFWorkbook) workbook).dispose();
        return false;
    }

    public String getFileExtension() {
        return workbook instanceof HSSFWorkbook ? "xls" : "xlsx";
    }

    /**
     * Returns the active sheet in this workbook.
     * <p>
     * The active sheet is the sheet which is displayed when a Microsoft Excel workbook is opened.
     * <p>
     * If this workbook does not contain any sheets then <i>Sheet1</i> will be created.
     * 
     * @return the active sheet in this workbook
     */
    public FSheet getOrCreateActiveSheet() {
        if (workbook.getNumberOfSheets() == 0)
            return getOrCreateSheet("Sheet1");
        return getOrCreateFromCache(workbook.getSheetAt(workbook.getActiveSheetIndex()));
    }

    /**
     * Returns the first sheet in this workbook.
     * <p>
     * If this workbook does not contain any sheets then <i>Sheet1</i> will be created.
     * 
     * @return the first sheet in this workbook
     */
    public FSheet getOrCreateFirstSheet() {
        return getOrCreateSheet(0);
    }

    /**
     * Returns the specified sheet in this workbook. If the number of sheets is less than the given index then new sheets
     * will be created until the specified index is reached.
     * 
     * @param index the 0-based index of the sheet to return
     * @return the specified sheet in this workbook
     */
    public FSheet getOrCreateSheet(final int index) {
        checkArgument(index >= 0, "sheet index < 0");
        IntStream.range(workbook.getNumberOfSheets() - 1, index).forEach(i -> workbook.createSheet("Sheet" + (i + 2)));
        return getOrCreateFromCache(workbook.getSheetAt(index));
    }

    /**
     * Returns the specified sheet in this workbook. If the sheet does not exist it will be created.
     * 
     * @param name the name of the sheet
     * @return the specified sheet in this workbook
     */
    public FSheet getOrCreateSheet(final String name) {
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final FSheet sheet = getSheet(name);
        if (sheet == null)
            return getOrCreateFromCache(workbook.createSheet(name));
        return sheet;
    }

    /**
     * Returns the specified sheet in this workbook or {@code null} if the number of sheets in the workbook is less than
     * {@code index}.
     * 
     * @param index the 0-based index of the sheet to return
     * @return the specified sheet in this workbook or {@code null}
     */
    public FSheet getSheet(final int index) {
        checkArgument(index >= 0, "sheet index < 0");
        if (workbook.getNumberOfSheets() - 1 < index)
            return null;
        return getOrCreateFromCache(workbook.getSheetAt(index));
    }

    /**
     * Returns the specified sheet in this workbook or {@code null} if the sheet does not exist.
     * 
     * @param name the name of the sheet
     * @return the specified sheet in this workbook or {@code null}
     */
    public FSheet getSheet(final String name) {
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        // We have to check if Workbook.getNumberNumberOfSheets == 0 to avoid a POI bug: in some cases when no sheets exist
        // Workbook.getSheetIndex(name) returns 0 and not -1
        final int index = workbook.getSheetIndex(name);
        return workbook.getNumberOfSheets() == 0 || index < 0 ? null : getOrCreateFromCache(workbook.getSheetAt(index));
    }

    @Override
    public Iterator<FSheet> iterator() {
        return Iterators.transform(workbook.iterator(), this::getOrCreateFromCache);
    }

    public FWorkbook open(final InputStream in, final Format format) throws IOException {
        checkNotNull(in, "in == null");
        checkNotNull(format, "format == null");
        return new FWorkbook(format == Format.BINARY ? new HSSFWorkbook(in) : format == Format.OFFICE_OPEN_XML ? new XSSFWorkbook(in) : new SXSSFWorkbook(new XSSFWorkbook(in)));
    }

    public FWorkbook open(final Path path) throws IOException {
        checkNotNull(path, "path == null");

        final String ext = MoreFiles.getFileExtension(path);
        final Format format;
        if (ext.equalsIgnoreCase("xls"))
            format = Format.BINARY;
        else if (ext.equalsIgnoreCase("xlsx"))
            format = Format.OFFICE_OPEN_XML;
        else
            throw new IllegalArgumentException("unknown extension: " + ext);

        try (final InputStream in = Files.newInputStream(path)) { // buffered?
            return open(in, format);
        }
    }

    public FWorkbook open(final Path path, final Format format) throws IOException {
        checkNotNull(path, "path == null");
        checkNotNull(format, "format == null");

        try (final InputStream in = Files.newInputStream(path)) { // do we want a buffered stream?
            return open(in, format);
        }
    }

    public FWorkbook removeSheet(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        final Workbook workbook = sheet.getWorkbook();
        workbook.removeSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
        return this;
    }

    public Path save(final Path path) throws IOException {
        return save(path, false);
    }

    public Path save(final Path path, final boolean close) throws IOException {
        checkNotNull(path, "path == null");
        try (final OutputStream out = Files.newOutputStream(path)) { // do we want a buffered stream?
            write(out, close);
        }
        return path;
    }

    public <T extends OutputStream> T write(final T out, final boolean close) throws IOException {
        checkNotNull(out, "out == null");

        IOException first = null;

        try {
            workbook.write(out);
        } catch (final IOException e) {
            first = e;
        }

        if (close) {
            try {
                workbook.close();
            } catch (final IOException e) {
                if (first == null)
                    first = e;
                else
                    first.addSuppressed(e);
            }
        }

        if (workbook instanceof SXSSFWorkbook && !dispose())
            logger.log(Level.WARNING, "SXSSFWorkbook.dispose() failed");

        if (first != null)
            throw first;

        return out;
    }

    private FSheet getOrCreateFromCache(final Sheet sheet) {
        try {
            return sheets.get(sheet, () -> new FSheet(this, sheet));
        } catch (ExecutionException e) {
            throw new AssertionError(); // cannot happen
        }
    }

}