package com.fenglex;


import com.fenglex.data.CellData;
import com.fenglex.data.SheetData;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;


public class Excel07Reader {

    private List<SheetData> sheets = new ArrayList<>(10);

    private SheetData currentSheet;
    private List<List<CellData>> lines;
    private List<CellData> line;

    public List<SheetData> getSheets() {
        return this.sheets;
    }

    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     * of parsing the Sheet XML, and outputs the contents
     * as a (basic) CSV.
     */
    private class SheetToCSV implements SheetContentsHandler {
        private boolean firstCellOfRow;
        private int currentRow = -1;
        private int currentCol = -1;

        private void outputMissingRows(int number) {
            for (int i = 0; i < number; i++) {
                for (int j = 0; j < minColumns; j++) {
                    //output.append(',');
                    line.add(new CellData(currentRow, j));
                }
                //output.append('\n');
            }
        }

        @Override
        public void startRow(int rowNum) {
            line = new ArrayList<>();
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum - currentRow - 1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
            // Ensure the minimum number of columns
            for (int i = currentCol; i < minColumns; i++) {
                //output.append(',');
                line.add(new CellData(currentRow, i));
            }
            //output.append('\n');
            lines.add(line);
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {
                //output.append(',');
                //line.add(new CellData(currentRow, currentCol));
            }
            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i = 0; i < missedCols; i++) {
                //output.append(',');
                line.add(new CellData(currentRow, currentCol));
            }
            currentCol = thisCol;

            // Number or string?
            line.add(new CellData(currentRow, currentCol, formattedValue));
            /*try {
                //noinspection ResultOfMethodCallIgnored
                Double.parseDouble(formattedValue);
                //output.append(formattedValue);
                line.add(new CellData(currentRow, currentCol, formattedValue));
            } catch (Exception e) {
                //output.append('"');
                //output.append(formattedValue);
                //output.append('"');
                line.add(new CellData(currentRow, currentCol, formattedValue));
            }*/
        }
    }


    ///////////////////////////////////////

    private final OPCPackage xlsxPackage;

    /**
     * Number of columns to read starting with leftmost
     */
    private final int minColumns;

    /**
     * Destination for data
     */
    // private final PrintStream output;

    /**
     * Creates a new XLSX -> CSV examples
     *
     * @param pkg        The XLSX package to process
     *                   //@param output     The PrintStream to output the CSV to
     * @param minColumns The minimum number of columns to output, or -1 for no minimum
     */
    public Excel07Reader(OPCPackage pkg, int minColumns) {
        this.xlsxPackage = pkg;
        // this.output = output;
        this.minColumns = minColumns;
        try {
            this.process();
        } catch (IOException | OpenXML4JException | SAXException e) {
            e.printStackTrace();
        }
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles           The table of styles that may be referenced by cells in the sheet
     * @param strings          The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.
     * @throws IOException  An IO exception from the parser,
     *                      possibly from a byte stream or character stream
     *                      supplied by the application.
     * @throws SAXException if parsing the XML data fails.
     */
    public void processSheet(
            Styles styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */

    public void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        while (iter.hasNext()) {
            SheetData sheet = new SheetData();
            sheets.add(sheet);
            currentSheet = sheet;
            try (InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                currentSheet.setName(sheetName);
                currentSheet.setIndex(index + 1);
                lines = new ArrayList<>();
                currentSheet.setData(lines);
                //this.output.println();
                //this.output.println(sheetName + " [index=" + index + "]:");
                processSheet(styles, strings, new SheetToCSV(), stream);
            }
            ++index;
        }
    }

    public static void main(String[] args) throws Exception {

        String file = "/Users/fenglex/Desktop/temp.xlsx";

        int minColumns = -1;
        try (OPCPackage p = OPCPackage.open(file, PackageAccess.READ)) {
            Excel07Reader reader = new Excel07Reader(p, minColumns);
            List<SheetData> sheets = reader.getSheets();
            //new CellReference()
            System.out.println(1);
        }


    }
}