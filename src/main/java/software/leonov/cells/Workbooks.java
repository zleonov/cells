package software.leonov.cells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Throwables;

/**
 * Static methods for working with {@link Workbook}s.
 * 
 * @author Zhenya Leonov
 */
final public class Workbooks {

    // private final static Logger logger = Logger.getLogger(Workbooks.class.getName());

    private Workbooks() {
    }

    /**
     * Specifies which Microsoft Excel format to use.
     */
    public enum Format {
        /**
         * The <i>xls</i> <a target="_blank" href= "http://en.wikipedia.org/wiki/Microsoft_Excel#Binary">Excel Binary File
         * Format</a> supported since Microsoft Office 2003.
         */
        BINARY_2003(65536, 256, "xls"),

        /**
         * The <i>xlsx</i> <a target="_blank" href= "https://en.wikipedia.org/wiki/Office_Open_XML">Office Open XML</a> format
         * supported starting with Microsoft Office 2007.
         */
        OFFICE_OPEN_XML(1048576, 16384, "xlsx"),

        /**
         * The streaming version of {@link #OFFICE_OPEN_XML Office Open XML} for writing large workbooks. See
         * {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new SXSSFWorkbook(XSSFWorkbook)} for more
         * information.
         * <p>
         * <b>Using this format is not recommended for developers who are making casual use of the API.</b>
         * <p>
         * <b>Note:</b> Allocated temporary files must be cleaned up explicitly by calling {@link Workbook#close()}.
         * <p>
         * See the {@link Workbooks#newWorkbook(Format, String...) newWorkbook(Format)} and {@link Workbooks#open(Path, Format)
         * open(Path, Format)} methods for further details.
         */
        STREAMING_OFFICE_OPEN_XML(1048576, 16384, "xlsx");

        private int    maxRows;
        private int    maxColumns;
        private String fileExtension;

        Format(final int maxRows, final int maxColumns, final String fileExtension) {
            this.maxRows       = maxRows;
            this.maxColumns    = maxColumns;
            this.fileExtension = fileExtension;
        }

        /**
         * Return the maximum number of columns (1-based) supported by this format.
         * 
         * @return the maximum number of columns (1-based) supported by this format
         */
        public int getMaxColNum() {
            return maxColumns;
        }

        /**
         * Return the maximum number of rows (1-based) supported by this format.
         * 
         * @return the maximum number of rows (1-based) supported by this format
         */
        public int getMaxRowNum() {
            return maxRows;
        }

        /**
         * Returns the file extension <i>xls</i> or <i>xlsx</i> corresponding to this format. The returned extension does not
         * include the leading dot character.
         * 
         * @return the file extension <i>xls</i> or <i>xlsx</i> corresponding to this format
         */
        public String getFileExtension() {
            return fileExtension;
        }

        /**
         * Returns the {@code Format} of the specified workbook.
         * 
         * @param workbook the specified workbook
         * @return the {@code Format} of the specified workbook
         */
        public static Format of(final Workbook workbook) {
            checkNotNull(workbook, "workbook == null");
            if (workbook instanceof HSSFWorkbook)
                return Format.BINARY_2003;
            else if (workbook instanceof XSSFWorkbook)
                return Format.OFFICE_OPEN_XML;
            else if (workbook instanceof SXSSFWorkbook)
                return Format.STREAMING_OFFICE_OPEN_XML;
            else
                throw new IllegalArgumentException("unkown workbook type: " + workbook.getClass().getSimpleName());
        }

    }

//    /**
//     * Dispose of temporary files backing an {@link SXSSFWorkbook} on disk. Calling this method will render the workbook
//     * unusable.
//     * <p>
//     * This method is a no-op for other {@link Workbook} implementations.
//     * 
//     * @param workbook the specified workbook
//     * @return {@code true} if the specified workbook is an {@code SXSSFWorkbook} and all temporary files were successfully
//     *         deleted
//     */
//    public static boolean dispose(final Workbook workbook) {
//        checkNotNull(workbook, "workbook == null");
//        if (workbook instanceof SXSSFWorkbook)
//            return ((SXSSFWorkbook) workbook).dispose();
//        return false;
//    }

    /**
     * Convenience method to get the active sheet from the specified workbook.
     * <p>
     * The active sheet is the sheet which is displayed when a Microsoft Excel workbook is opened.
     * <p>
     * If the workbook does not contain any sheets then <i>Sheet1</i> will be created.
     * 
     * @param workbook the specified workbook
     * @return the active sheet from the specified workbook
     */
    public static Sheet getOrCreateActiveSheet(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        if (workbook.getNumberOfSheets() == 0)
            return getOrCreateSheet(workbook, "Sheet1");
        return workbook.getSheetAt(workbook.getActiveSheetIndex());
    }

    /**
     * Convenience method to get the first sheet from the specified workbook.
     * <p>
     * If the workbook does not contain any sheets then <i>Sheet1</i> will be created.
     * 
     * @param workbook the specified workbook
     * @return the first sheet in the specified workbook
     */
    public static Sheet getOrCreateFirstSheet(final Workbook workbook) {
        checkNotNull(workbook, "workbook == null");
        return getOrCreateSheet(workbook, 0);
    }

    /**
     * Returns the specified sheet from the workbook. If the number of sheets in this workbook is less than the given index
     * then new sheets will be created until the specified index is reached.
     * 
     * @param workbook the workbook
     * @param index    the 0-based index of the sheet to return
     * @return the specified sheet from the workbook
     */
    public static Sheet getOrCreateSheet(final Workbook workbook, final int index) {
        checkNotNull(workbook, "workbook == null");
        checkArgument(index >= 0, "sheet index < 0");
        IntStream.range(workbook.getNumberOfSheets() - 1, index).forEach(i -> workbook.createSheet("Sheet" + (i + 2)));
        return workbook.getSheetAt(index);
    }

    /**
     * Returns the specified sheet from the workbook. If the sheet does not exist it will be created.
     * 
     * @param workbook the workbook
     * @param name     the name of the sheet
     * @return the specified sheet from the workbook
     */
    public static Sheet getOrCreateSheet(final Workbook workbook, final String name) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);
        final Sheet sheet = getSheet(workbook, name);
        if (sheet == null)
            return workbook.createSheet(name);
        return sheet;
    }

    /**
     * Returns the specified sheet from the workbook or {@code null} if the number of sheets in the workbook is less than
     * {@code index}.
     * 
     * @param workbook the specified workbook
     * @param index    the 0-based index sheet index
     * @return the specified sheet from the workbook or {@code null}
     */
    public static Sheet getSheet(final Workbook workbook, final int index) {
        checkNotNull(workbook, "workbook == null");
        checkArgument(index >= 0, "sheet index < 0");
        if (workbook.getNumberOfSheets() - 1 < index)
            return null;
        return workbook.getSheetAt(index);
    }

    /**
     * Returns the specified sheet from the workbook or {@code null} if the sheet does not exist.
     * 
     * @param workbook the workbook
     * @param name     the name of the sheet
     * @return the specified sheet from the workbook or {@code null}
     */
    public static Sheet getSheet(final Workbook workbook, final String name) {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(name, "name == null");
        WorkbookUtil.validateSheetName(name);

        final int index = workbook.getSheetIndex(name);

        /*
         * We have to check if Workbook.getNumberOfSheets == 0 to avoid a bug where some implementations
         * Workbook.getSheetIndex(name) returns 0 and not -1 when no sheet exist
         */
        return workbook.getNumberOfSheets() == 0 || index < 0 ? null : workbook.getSheetAt(index);
    }

    /**
     * Creates a new {@code Workbook} with an empty sheet.
     * <p>
     * Workbooks can be created in the classic {@link Format#BINARY_2003 Excel Binary File Format} {@code xls} format or the
     * {@link Format#OFFICE_OPEN_XML Office Open XML} {@code xlsx} format.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information. Consider creating an {@link SXSSFWorkbook} manually if explicit
     * control over its behavior is desired.
     * 
     * @param format specifies which workbook format to use
     * @return a new {@code Workbook} with an empty sheet
     */
//    public static Workbook newWorkbook(final Format format) {
//        checkNotNull(format, "format == null");
//        return newWorkbook(format, 1);
//    }

    /**
     * Creates a new {@code Workbook} and adds the specified number of empty sheets.
     * <p>
     * Workbooks can be created in the classic {@link Format#BINARY_2003 Excel Binary File Format} {@code xls} format or the
     * {@link Format#OFFICE_OPEN_XML Office Open XML} {@code xlsx} format.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information. Consider creating an {@link SXSSFWorkbook} manually if explicit
     * control over its behavior is desired.
     * 
     * @param format  specifies which workbook format to use
     * @param nsheets the number of sheets to create in the workbook
     * @return a new {@code Workbook} containing the specified number of empty sheets named <i>Sheet1</i>, <i>Sheet2</i>,
     *         <i>Sheet3</i>, etc...
     */
    public static Workbook newWorkbook(final Format format, final int nsheets) {
        checkArgument(nsheets >= 0, "nsheets < 0");

        final String[] sheets = (String[]) IntStream.range(1, nsheets + 1).mapToObj(i -> "Sheet" + i).collect(Collectors.toList()).toArray();

        return newWorkbook(format, sheets);
    }

    /**
     * Creates a new {@code Workbook} and with the specified named sheets.
     * <p>
     * Workbooks can be created in the classic {@link Format#BINARY_2003 Excel Binary File Format} {@code xls} format or the
     * {@link Format#OFFICE_OPEN_XML Office Open XML} {@code xlsx} format.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information. Consider creating an {@link SXSSFWorkbook} manually if explicit
     * control over its behavior is desired.
     * 
     * @param format specifies which workbook format to use
     * @param sheets any additional sheets
     * @return a new {@code Workbook} containing the specified named sheets
     */
    public static Workbook newWorkbook(final Format format, final String... sheets) {
        checkNotNull(format, "format == null");
        checkNotNull(sheets, "sheets == null");

        final Workbook workbook = format == Format.BINARY_2003 ? new HSSFWorkbook() : format == Format.OFFICE_OPEN_XML ? new XSSFWorkbook() : new SXSSFWorkbook();

        for (final String name : sheets) {
            WorkbookUtil.validateSheetName(name);
            workbook.createSheet(name);
        }

        if (sheets.length > 0)
            workbook.setActiveSheet(0);

        return workbook;
    }

    /**
     * Opens a workbook from the specified input stream. Does not close the stream.
     * <p>
     * This method will automatically determine the format as either {@link Format#BINARY_2003} or
     * {@link Format#OFFICE_OPEN_XML} based on the stream content.
     * 
     * @param in the specified input stream
     * @return a new workbook loaded from the specified input stream
     * @throws IOException if an I/O error occurs
     * @see HSSFWorkbook
     * @see XSSFWorkbook
     * @see SXSSFWorkbook
     * @see WorkbookFactory
     */
    public static Workbook open(final InputStream in) throws IOException {
        checkNotNull(in, "in == null");
        return WorkbookFactory.create(in);
    }

    /**
     * Opens a workbook from the specified input stream. Does not close the stream.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information. Consider creating an {@link SXSSFWorkbook} manually if explicit
     * control over its behavior is desired.
     * 
     * @param in the specified input stream
     * @return a new workbook loaded from the specified input stream
     * @param format specifies which workbook format to use
     * @throws IOException if an I/O error occurs
     * @see HSSFWorkbook
     * @see XSSFWorkbook
     * @see SXSSFWorkbook
     * @see WorkbookFactory
     */
    public static Workbook open(final InputStream in, final Format format) throws IOException {
        checkNotNull(in, "in == null");
        checkNotNull(format, "format == null");

        return format == Format.BINARY_2003 ? new HSSFWorkbook(in) : format == Format.OFFICE_OPEN_XML ? new XSSFWorkbook(in) : new SXSSFWorkbook(new XSSFWorkbook(in));
    }

    /**
     * Opens a workbook from the specified path.
     * <p>
     * This method will automatically determine the format as either {@link Format#BINARY_2003} or
     * {@link Format#OFFICE_OPEN_XML} based on the file content.
     * 
     * @param path the path to load
     * @return a new workbook object loaded from the specified path
     * @throws IOException if an I/O error occurs
     * @see HSSFWorkbook
     * @see XSSFWorkbook
     * @see SXSSFWorkbook
     * @see WorkbookFactory
     */
    public static Workbook open(final Path path) throws IOException {
        checkNotNull(path, "path == null");

        try (final InputStream in = new BufferedInputStream(Files.newInputStream(path))) {
            return WorkbookFactory.create(in);
        }
    }

    /**
     * Opens a workbook from the specified path.
     * <p>
     * {@link Format#STREAMING_OFFICE_OPEN_XML Streaming Office Open XML} workbooks will be created using the default
     * settings. See {@link SXSSFWorkbook} and {@link SXSSFWorkbook#SXSSFWorkbook(XSSFWorkbook) new
     * SXSSFWorkbook(XSSFWorkbook)} for more information. Consider creating an {@link SXSSFWorkbook} manually if explicit
     * control over its behavior is desired.
     * 
     * @param path   the path to load
     * @param format specifies which workbook format to use
     * @return a new workbook loaded from the specified file
     * @throws IOException if an I/O error occurs
     * @see HSSFWorkbook
     * @see XSSFWorkbook
     * @see SXSSFWorkbook
     * @see WorkbookFactory
     */
    public static Workbook open(final Path path, final Format format) throws IOException {
        checkNotNull(path, "path == null");
        checkNotNull(format, "format == null");

        try (final InputStream in = new BufferedInputStream(Files.newInputStream(path))) {
            return open(in, format);
        }
    }

//    /**
//     * Writes the given workbook to a file in the default temporary-file directory.
//     * 
//     * @param workbook the given workbook
//     * @return the {@code File} object representing the temporary-file
//     * @throws IOException if an I/O error occurs
//     */
//    public static File saveAsTemp(final Workbook workbook) throws IOException {
//        checkNotNull(workbook, "workbook == null");
//        final String suffix;
//        if (workbook instanceof HSSFWorkbook)
//            suffix = ".xls";
//        else
//            suffix = ".xlsx";
//        final File path = File.createTempFile("tmp", suffix);
//        Workbooks.saveAs(workbook, path);
//        return path;
//    }
//
//    /**
//     * Writes the given workbook to a {@code ByteArrayOutputStream} and returns the contents as a byte array.
//     * 
//     * @param workbook the given workbook
//     * @return the contents of the workbook as a byte array
//     * @throws IOException if an I/O error occurs
//     */
//    public static byte[] toByteArray(final Workbook workbook) throws IOException {
//        checkNotNull(workbook, "workbook == null");
//        return saveAs(workbook, new ByteArrayOutputStream()).toByteArray();
//    }

    /**
     * Removes a sheet from this workbook.
     * 
     * @param sheet the sheet to remove
     * @return the workbook where the sheet was located
     */
    public static Workbook removeSheet(final Sheet sheet) {
        checkNotNull(sheet, "sheet == null");
        final Workbook workbook = sheet.getWorkbook();
        workbook.removeSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
        return workbook;
    }

    /**
     * Writes the given workbook to the specified path and {@link Workbook#close() closes} the workbook.
     * 
     * @param workbook the given workbook
     * @param path     the specified path
     * @throws IOException if an I/O error occurs
     * @return the specified file
     */
    public static Path save(final Workbook workbook, final Path path) throws IOException {
        return save(workbook, path, true);
    }

    /**
     * Writes the given workbook to the specified path.
     * 
     * @param workbook the given workbook
     * @param path     the specified path
     * @param close    whether or not to {@link Workbook#close() close} the workbook
     * @throws IOException if an I/O error occurs
     * @return the specified file
     */
    public static Path save(final Workbook workbook, final Path path, final boolean close) throws IOException {
        checkNotNull(workbook, "workbook == null");
        checkNotNull(path, "path == null");
        try (final OutputStream out = new BufferedOutputStream(Files.newOutputStream(path))) {
            write(workbook, out, close);
        }
        return path;
    }

    /**
     * Writes the given workbook to the specified output stream. Does not close the stream.
     * 
     * @param <T>      the type of output stream
     * @param workbook the given workbook
     * @param out      the specified output stream
     * @param close    whether or not to {@link Workbook#close() close} the workbook
     * @throws IOException if an I/O error occurs
     * @return the specified output stream
     */
    public static <T extends OutputStream> T write(final Workbook workbook, final T out, final boolean close) throws IOException {
        checkNotNull(out, "out == null");

        Throwable first = null;

        try {
            workbook.write(out);
        } catch (final Throwable t) {
            first = t;
        }

        if (close)
            try {
                workbook.close();
            } catch (final Throwable t) {
                if (first == null)
                    first = t;
                else
                    first.addSuppressed(t);
            }

//        if (close && workbook instanceof SXSSFWorkbook && !dispose(workbook))
//            logger.log(Level.WARNING, "SXSSFWorkbook.dispose() failed");

        if (first != null)
            Throwables.propagateIfPossible(first, IOException.class);

        return out;
    }

//    /**
//     * Returns the {@link Format} of the specified workbook.
//     * 
//     * @param workbook the specified workbook
//     * @return the {@link Format} of the specified workbook
//     */
//    public static Format getFormatOf(final Workbook workbook) {
//        checkNotNull(workbook, "workbook == null");
//        if (workbook instanceof HSSFWorkbook)
//            return Format.BINARY_2003;
//        else if (workbook instanceof XSSFWorkbook)
//            return Format.OFFICE_OPEN_XML;
//        else if (workbook instanceof SXSSFWorkbook)
//            return Format.STREAMING_OFFICE_OPEN_XML;
//        else
//            throw new IllegalArgumentException("unkown workbook type: " + workbook.getClass().getSimpleName());
//    }
//
//    /**
//     * Returns the file extension <i>xls</i> or <i>xlsx</i> corresponding the specified workbook. The returned extension
//     * does not include the leading dot character.
//     * 
//     * @param workbook the specified workbook
//     * @return the file extension <i>xls</i> or <i>xlsx</i> corresponding the specified workbook
//     */
//    public static String getFileExtension(final Workbook workbook) {
//        checkNotNull(workbook, "workbook == null");
//        return getFormatOf(workbook).fileExtension();
//    }

}