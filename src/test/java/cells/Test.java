package cells;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import software.leonov.cells.Cells;
import software.leonov.cells.Rows;
import software.leonov.cells.Sheets;
import software.leonov.cells.Workbooks;
import software.leonov.cells.Workbooks.Format;
import software.leonov.cells.util.FontBuilder;
import software.leonov.cells.util.StyleBuilder;

public class Test {

    public static void main(String[] args) throws IOException {

        final Workbook book = Workbooks.newWorkbook(Format.OFFICE_OPEN_XML);

        final Path path = Files.createTempFile("test", ".xlsx");

        final CellStyle style = new StyleBuilder().setSolidFillColor(IndexedColors.GREY_25_PERCENT).setFont(new FontBuilder().setBold(true).create(book)).create(book);

        final Sheet sheet = Workbooks.getOrCreateActiveSheet(book);

        final Row row = Sheets.createNextRow(sheet);

        final Cell cell = Rows.getOrCreateCell(row, "A");

        Cells.setValue(cell, "A1");
        Cells.setStyle(cell, style);

        Sheets.setColumnStyle(sheet, 2, style, true);
        final Row  row2  = Sheets.getOrCreateRow(sheet, 5);
        final Cell cell2 = Rows.getOrCreateCell(row2, "C");
        Cells.setValue(cell2, "try2");
        Cells.setValue(cell2, null);

        Workbooks.save(book, path, true);

        Desktop.getDesktop().open(path.toFile());

    }

}
