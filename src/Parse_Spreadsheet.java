
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Parse_Spreadsheet {



    public static ArrayList<String> returnNames() throws IOException {
        FileInputStream file;
        ArrayList<String> emails = new ArrayList<>();
        ArrayList<String> names = new ArrayList<>();

        try {

            file = new FileInputStream("TestFile.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            break;
                        case STRING:
                            if (cell.getStringCellValue().contains("@")) {
                                emails.add(cell.getStringCellValue());
                            } else {
                                names.add(cell.getStringCellValue());
                            }
                            break;
                    }

                }

            }
            file.close();
            return names;

        } catch (IOException e) {
            e.printStackTrace();
        }


        return names;
    }

    public static ArrayList<String> returnEmails() throws IOException {
        FileInputStream file;
        ArrayList<String> emails = new ArrayList<>();
        ArrayList<String> names = new ArrayList<>();

        try {

            file = new FileInputStream("TestFile.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            break;
                        case STRING:
                            if (cell.getStringCellValue().contains("@")) {
                                emails.add(cell.getStringCellValue());
                            } else {
                                names.add(cell.getStringCellValue());
                            }
                            break;
                    }

                }

            }
            file.close();
            return names;

        } catch (IOException e) {
            e.printStackTrace();
        }


        return emails;
    }
}

