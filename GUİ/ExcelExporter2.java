package GUİ;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.table.DefaultTableModel;
import java.io.*;
import java.util.Iterator;

public class ExcelExporter2 {

    private static final String FILE_NAME = "faturalar.xlsx";

    public static void writeInvoice(String tarih, String baslik, String vergi, double tutar) throws IOException {
        File file = new File(FILE_NAME);
        Workbook workbook;
        Sheet sheet;

        if (file.exists()) {
            FileInputStream fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);
            sheet = workbook.getSheetAt(0);
            fis.close();
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Faturalar");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Tarih");
            header.createCell(1).setCellValue("Başlık");
            header.createCell(2).setCellValue("Vergi (%)");
            header.createCell(3).setCellValue("Tutar (€)");
        }

        int rowNum = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(tarih);
        row.createCell(1).setCellValue(baslik);
        row.createCell(2).setCellValue(vergi);
        row.createCell(3).setCellValue(tutar);

        FileOutputStream fos = new FileOutputStream(FILE_NAME);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }

    public static void loadInvoices(DefaultTableModel model) {
        File file = new File(FILE_NAME);
        if (!file.exists()) return;

        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            if (rowIterator.hasNext()) rowIterator.next(); // skip header

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String tarih = row.getCell(0).getStringCellValue();
                String baslik = row.getCell(1).getStringCellValue();
                String vergi = row.getCell(2).getStringCellValue();
                double tutar = row.getCell(3).getNumericCellValue();

                model.addRow(new Object[]{tarih, baslik, vergi, String.format("%.2f", tutar)});
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
