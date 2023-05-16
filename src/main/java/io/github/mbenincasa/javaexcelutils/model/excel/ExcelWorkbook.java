package io.github.mbenincasa.javaexcelutils.model.excel;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import io.github.mbenincasa.javaexcelutils.exceptions.OpenWorkbookException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetAlreadyExistsException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetNotFoundException;
import lombok.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import io.github.mbenincasa.javaexcelutils.tools.ExcelUtility;

import java.io.*;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * {@code ExcelWorkbook} is the {@code Workbook} wrapper class of the Apache POI library
 * @author Mirko Benincasa
 * @since 0.3.0
 */
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Getter
@EqualsAndHashCode
@Builder(access = AccessLevel.PRIVATE)
public class ExcelWorkbook {

    /**
     * This object refers to the Apache POI Library {@code Workbook}
     */
    private Workbook workbook;

    /**
     * This constructor creates a new workbook based on the extension
     * @param extension The Excel extension which will determine with which Excel version the Workbook will be created
     */
    private ExcelWorkbook(Extension extension) {
        switch (extension) {
            case XLS -> this.workbook = new HSSFWorkbook();
            case XLSX -> this.workbook = new XSSFWorkbook();
        }
    }

    /**
     * This constructor opens a Workbook from the InputStream
     * @param inputStream The InputStream of an existing Excel file
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    private ExcelWorkbook(InputStream inputStream) throws OpenWorkbookException {
        try {
            this.workbook = new XSSFWorkbook(inputStream);
        } catch (OfficeXmlFileException | OLE2NotOfficeXmlFileException | IOException e) {
            try {
                this.workbook = new HSSFWorkbook(inputStream);
            } catch (IOException ex) {
                throw new OpenWorkbookException("The workbook could not be opened", ex);
            }
        }
    }

    /**
     * Get to an ExcelWorkbook instance from Apache POI Workbook
     * @param workbook The Workbook instance to wrap
     * @return The ExcelWorkbook instance
     * @since 0.5.0
     */
    public static ExcelWorkbook of(Workbook workbook) {
        return ExcelWorkbook.builder()
                .workbook(workbook)
                .build();
    }

    /**
     * Opens the workbook
     * @param file An Excel file
     * @return An ExcelWorkBook that is represented in the Excel file
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static ExcelWorkbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        ExcelUtility.checkExcelExtension(file.getName());

        /* Open file input stream */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbook excelWorkbook = open(fileInputStream);

        /* Close the stream before return */
        fileInputStream.close();
        return excelWorkbook;
    }

    /**
     * Opens the workbook
     * @param inputStream The {@code InputStream} of the Excel file
     * @return An ExcelWorkBook that is represented in the Excel file
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static ExcelWorkbook open(InputStream inputStream) throws OpenWorkbookException {
        return new ExcelWorkbook(inputStream);
    }

    /**
     * Create a new workbook<p>
     * If not specified the XLSX extension will be used
     * @return A ExcelWorkbook
     */
    public static ExcelWorkbook create() {
        return create(Extension.XLSX);
    }

    /**
     * Create a new workbook
     * @param extension The extension of the file. Provide the extension of an Excel file
     * @return A ExcelWorkbook
     * @throws ExtensionNotValidException If the extension does not belong to an Excel file
     */
    public static ExcelWorkbook create(String extension) throws ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return create(Extension.getExcelExtension(extension));
    }

    /**
     * Create a new workbook
     * @param extension The extension of the file. Select an extension with {@code type} EXCEL
     * @return A ExcelWorkbook
     */
    public static ExcelWorkbook create(Extension extension) {
        return new ExcelWorkbook(extension);
    }

    /**
     * Close a workbook
     * @throws IOException If an I/O error has occurred
     */
    public void close() throws IOException {
        this.workbook.close();
    }

    /**
     * Close a workbook
     * @param inputStream The {@code InputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public void close(InputStream inputStream) throws IOException {
        this.workbook.close();
        inputStream.close();
    }

    /**
     * Close a workbook
     * @param outputStream The {@code OutputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public void close(OutputStream outputStream) throws IOException {
        this.workbook.close();
        outputStream.close();
    }

    /**
     * Close a workbook
     * @param outputStream The {@code OutputStream} to close
     * @param inputStream The {@code InputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public void close(OutputStream outputStream, InputStream inputStream) throws IOException {
        this.workbook.close();
        inputStream.close();
        outputStream.close();
    }

    /**
     * Close a workbook
     * @param writer The {@code CSVWriter} to close
     * @throws IOException If an I/O error has occurred
     */
    public void close(CSVWriter writer) throws IOException {
        this.workbook.close();
        writer.close();
    }

    /**
     * Close a workbook
     * @param outputStream The {@code OutputStream} to close
     * @param reader The {@code CSVReader} to close
     * @throws IOException If an I/O error has occurred
     */
    public void close(OutputStream outputStream, CSVReader reader) throws IOException {
        this.workbook.close();
        outputStream.close();
        reader.close();
    }

    /**
     * The amount of Sheets in the Workbook
     * @return The number of Sheets present
     */
    public Integer countSheets() {
        return this.workbook.getNumberOfSheets();
    }

    /**
     * The list of Sheets related to the Workbook
     * @return A list of Sheets
     */
    public List<ExcelSheet> getSheets() {
        List<ExcelSheet> excelSheets = new LinkedList<>();
        for (Sheet sheet : this.workbook) {
            excelSheets.add(ExcelSheet.of(sheet));
        }
        return excelSheets;
    }

    /**
     * Create a new Sheet inside the Workbook
     * @return The newly created Sheet
     */
    public ExcelSheet createSheet() {
        Sheet sheet = this.workbook.createSheet();
        return ExcelSheet.of(sheet);
    }

    /**
     * Create a new Sheet inside the Workbook
     * @param sheetName The name of the sheet to create
     * @return The newly created Sheet
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    public ExcelSheet createSheet(String sheetName) throws SheetAlreadyExistsException {
        try {
            Sheet sheet = this.workbook.createSheet(sheetName);
            return ExcelSheet.of(sheet);
        } catch (IllegalArgumentException ex) {
            throw new SheetAlreadyExistsException(ex.getMessage());
        }
    }

    /**
     * Retrieve the Sheet with index 0
     * @return The Sheet requested
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public ExcelSheet getSheet() throws SheetNotFoundException {
        return this.getSheet(0);
    }

    /**
     * Retrieve the Sheet with the requested name
     * @param index The index in the workbook
     * @return The Sheet requested
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public ExcelSheet getSheet(Integer index) throws SheetNotFoundException {
        List<ExcelSheet> excelSheets = this.getSheets();
        for (ExcelSheet excelSheet : excelSheets) {
            if (Objects.equals(excelSheet.getIndex(), index))
                return excelSheet;
        }

        throw new SheetNotFoundException("No sheet was found in the index: " + index);
    }

    /**
     * Retrieve the Sheet with the requested index
     * @param sheetName The name of the sheet
     * @return The Sheet requested
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public ExcelSheet getSheet(String sheetName) throws SheetNotFoundException {
        List<ExcelSheet> excelSheets = this.getSheets();
        for (ExcelSheet excelSheet : excelSheets) {
            if (excelSheet.getName().equals(sheetName))
                return excelSheet;
        }

        throw new SheetNotFoundException("No sheet was found with the name: " + sheetName);
    }

    /**
     * Retrieve the Sheet with the required name otherwise create it
     * @param sheetName The name of the sheet
     * @return The Sheet requested
     */
    @SneakyThrows
    public ExcelSheet getSheetOrCreate(String sheetName) {
        try {
            return this.getSheet(sheetName);
        } catch (SheetNotFoundException e) {
            return this.createSheet(sheetName);
        }
    }

    /**
     * Remove the Sheet
     * @param index The index of the Sheet in the workbook that will be removed
     * @since 0.4.1
     */
    public void removeSheet(Integer index) {
        this.workbook.removeSheetAt(index);
    }

    /**
     * Check if the sheet is present
     * @param sheetName The name of the sheet
     * @return {@code true} if is present
     */
    public Boolean isSheetPresent(String sheetName) {
        List<ExcelSheet> excelSheets = this.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> s.getName().equals(sheetName)).findAny();
        return excelSheet.isPresent();
    }

    /**
     * Check if the sheet is present
     * @param index The index in the workbook
     * @return {@code true} if is present
     */
    public Boolean isSheetPresent(Integer index) {
        List<ExcelSheet> excelSheets = this.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> Objects.equals(s.getIndex(), index)).findAny();
        return excelSheet.isPresent();
    }

    /**
     * Check if the sheet is not present
     * @param sheetName The name of the sheet
     * @return {@code true} if is not present
     */
    public Boolean isSheetNull(String sheetName) {
        List<ExcelSheet> excelSheets = this.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> s.getName().equals(sheetName)).findAny();
        return excelSheet.isEmpty();
    }

    /**
     * Check if the sheet is not present
     * @param index The index in the workbook
     * @return {@code true} if is not present
     */
    public Boolean isSheetNull(Integer index) {
        List<ExcelSheet> excelSheets = this.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> Objects.equals(s.getIndex(), index)).findAny();
        return excelSheet.isEmpty();
    }

    /**
     * Create a new FormulaEvaluator
     * @return A FormulaEvaluator
     */
    public FormulaEvaluator getFormulaEvaluator() {
        return this.workbook.getCreationHelper().createFormulaEvaluator();
    }


    /**
     * Writes the OutputStream to the Workbook and then closes them
     * @param outputStream The {@code OutputStream} to close
     * @throws IOException If an I/O error has occurred
     * @since 0.4.0
     */
    public void writeAndClose(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
        this.close(outputStream);
    }
}
