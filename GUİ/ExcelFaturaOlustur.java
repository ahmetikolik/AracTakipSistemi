package GUİ;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelFaturaOlustur {

    public static void olustur(
            String taxiNr,
            String rechnungNr,
            String leistungszeitraum,
            String tarihStr,
            String iban,
            double net7, double tax7,
            double net19, double tax19,
            double altNet19, double altTax19
    ) {
        double brut7 = net7 + tax7;
        double brut19 = net19 + tax19;
        double altBrut = altNet19 + altTax19;
        double total = brut7 + brut19 + altBrut;

        try {
            File dosya = new File("faturalar.xlsx");
            Workbook workbook;

            if (dosya.exists()) {
                FileInputStream fis = new FileInputStream(dosya);
                workbook = new XSSFWorkbook(fis);
                fis.close();
            } else {
                workbook = new XSSFWorkbook();
            }

            String sheetAdi = "Fatura_" + new SimpleDateFormat("ddMMyy_HHmmss").format(new Date());
            Sheet sheet = workbook.createSheet(sheetAdi);

            Font bold = workbook.createFont();
            bold.setBold(true);
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(bold);

            // Üst bilgiler
            Row row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue("Taxiunternehmen");
            row1.getCell(0).setCellStyle(boldStyle);
            row1.createCell(1).setCellValue("TAXIone GmbH Trierer Str. 94 52078 Aachen");

            sheet.createRow(2).createCell(0).setCellValue("Das obige Taxiunternehmen stellt den Aachener Autodroschken Bendelstr. 28-32 52062 Aachen");
            sheet.createRow(3).createCell(0).setCellValue("in Rechnung für durchgeführte Taxifahrten, Krankenfahrten - laut beigefügter Einzelaufstellung");

            sheet.createRow(5).createCell(0).setCellValue("Taxi Nr:");
            sheet.getRow(5).createCell(1).setCellValue(taxiNr);

            sheet.createRow(6).createCell(0).setCellValue("Rechnungs.Nr:");
            sheet.getRow(6).createCell(1).setCellValue(rechnungNr);

            sheet.createRow(7).createCell(0).setCellValue("Kalenderwoche :");
            sheet.getRow(7).createCell(1).setCellValue("0");

            sheet.createRow(8).createCell(0).setCellValue("Leistungszeitraum :");
            sheet.getRow(8).createCell(1).setCellValue(leistungszeitraum);

            sheet.createRow(9).createCell(0).setCellValue("Steuer Nr:");
            sheet.getRow(9).createCell(1).setCellValue("201/5993/4452");

            sheet.getRow(5).createCell(4).setCellValue("Aachener Autodroschken Vereinigung");
            sheet.getRow(6).createCell(4).setCellValue("Bendelstr. 28 - 32");
            sheet.getRow(7).createCell(4).setCellValue("52062 Aachen");
            sheet.getRow(9).createCell(4).setCellValue("Datum.");
            sheet.getRow(9).createCell(5).setCellValue(tarihStr);

            // Fatura başlığı
            Row r12 = sheet.createRow(12);
            r12.createCell(0).setCellValue("Rechnung");
            r12.getCell(0).setCellStyle(boldStyle);

            // Tablo başlıkları
            Row r14 = sheet.createRow(14);
            r14.createCell(1).setCellValue("Nettobetrag");
            r14.createCell(2).setCellValue("Steuer");
            r14.createCell(3).setCellValue("Bruttobetrag");

            sheet.createRow(15).createCell(0).setCellValue("Beträge mit 0% Mwst");
            sheet.getRow(15).createCell(2).setCellValue("Auslandsfahrten");
            sheet.getRow(15).createCell(3).setCellValue("0,00 €");

            Row r16 = sheet.createRow(16);
            r16.createCell(0).setCellValue("Beträge zu 7 % Mwst");
            r16.createCell(1).setCellValue(formatEuro(net7));
            r16.createCell(2).setCellValue(formatEuro(tax7));
            r16.createCell(3).setCellValue(formatEuro(brut7));

            Row r17 = sheet.createRow(17);
            r17.createCell(0).setCellValue("Beträge zu 19 %");
            r17.createCell(1).setCellValue(formatEuro(net19));
            r17.createCell(2).setCellValue(formatEuro(tax19));
            r17.createCell(3).setCellValue(formatEuro(brut19));

            Row r18 = sheet.createRow(18);
            r18.createCell(0).setCellValue("ALT/ASA/RED zu 19 %");
            r18.createCell(1).setCellValue(formatEuro(altNet19));
            r18.createCell(2).setCellValue(formatEuro(altTax19));
            r18.createCell(3).setCellValue(formatEuro(altBrut));

            Row r21 = sheet.createRow(21);
            r21.createCell(2).setCellValue("Gesamtbetrag  in Euro");
            r21.createCell(4).setCellValue(formatEuro(total));

            sheet.createRow(23).createCell(2).setCellValue("IBAN");
            sheet.getRow(23).createCell(4).setCellValue(iban);

            sheet.createRow(25).createCell(0).setCellValue("Diese Rechnung wurde maschinell erstellt, sie gilt ohne Unterschrift.");
            sheet.createRow(26).createCell(0).setCellValue("Bitte überweisen Sie den Betrag ohne Abzug auf mein obiges Konto !");

            // Yaz
            FileOutputStream fos = new FileOutputStream(dosya);
            workbook.write(fos);
            fos.close();
            workbook.close();

            System.out.println("✅ Yeni fatura eklendi: " + sheetAdi);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String formatEuro(double val) {
        return String.format("%.2f €", val);
    }
}
