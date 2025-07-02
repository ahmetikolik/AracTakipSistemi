package GUİ;

//POI - Excel işlemleri
import GUİ.ExcelExporter;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
//Swing ve GUI
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SpinnerDateModel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableCellRenderer;
//Ekstra (JTextComponent, EventListener vs.)
import javax.swing.text.JTextComponent;
import java.util.Calendar;
import java.util.Date;
//AWT
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
//IO
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
//Veri yapıları
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

//Tarih bileşeni (JDateChooser)
import com.toedter.calendar.JDateChooser;
public class anamenü {
	private String seciliExcelYolu = "surusler.xlsx";
    private JFrame frame;
    private JTable table;
    private DefaultTableModel tableModel;
    private JPanel formPanel;
    private JPanel filtrePanel; 

    public anamenü() {
    	
        initialize();
        try {
        	ExcelExporter.exportToExcel(tableModel,"surusler.xlsx");
        	
      
        } catch (Exception e) {
            JOptionPane.showMessageDialog(frame, "Excel dosyası okunamadı: " + e.getMessage());
        }
    }
  private void updateBakimDurumu(String plaka, int yeniDurum) {
    try {
        List<String> lines = new ArrayList<>();
        java.nio.file.Path path = java.nio.file.Paths.get("plakalar.csv");
        java.nio.file.Files.lines(path).forEach(line -> {
            String[] parts = line.split(",");
            if (parts.length >= 5 && parts[0].equals(plaka)) {
                // Güncel format: plaka, basKm, sonKm, bakımDurumu, muayeneAyi
                String yeniSatir = parts[0] + "," + parts[1] + "," + parts[2] + "," + yeniDurum + "," + parts[4];
                lines.add(yeniSatir);
            } else {
                lines.add(line);
            }
        });
        java.nio.file.Files.write(path, lines);
    } catch (IOException e) {
        JOptionPane.showMessageDialog(frame, "Bakım güncellenemedi: " + e.getMessage());
    }
}


    private void setFormPanel(JPanel newPanel) {
        formPanel.removeAll();
        formPanel.add(newPanel);
        formPanel.revalidate();
        formPanel.repaint();

        // Bu satırı EKLE: JComboBox gibi bileşenleri güncellemesi için
        frame.revalidate();
        frame.repaint();

        yenidenTabloyuYukle(); // Eğer tablo da güncellenecekse
    }
    private int hesaplaToplamGelir(String plaka) {
        int toplam = 0;
        try (FileInputStream fis = new FileInputStream("surusler.xlsx")) {
            Workbook wb = new XSSFWorkbook(fis);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell plakaCell = row.getCell(1); // Plaka sütunu
                    Cell gelirCell = row.getCell(10); // Umsatz sütunu

                    if (plakaCell != null && gelirCell != null) {
                        String plakaText = (plakaCell.getCellType() == CellType.STRING)
                                ? plakaCell.getStringCellValue().trim()
                                : String.valueOf((int) plakaCell.getNumericCellValue());

                        if (plaka.equals(plakaText)) {
                            if (gelirCell.getCellType() == CellType.NUMERIC) {
                                toplam += (int) gelirCell.getNumericCellValue();
                            } else if (gelirCell.getCellType() == CellType.STRING) {
                                String raw = gelirCell.getStringCellValue().trim().replace(",", ".");
                                try {
                                    toplam += (int) Double.parseDouble(raw);
                                } catch (NumberFormatException ex) {
                                    System.err.println("Umsatz değeri çözümlenemedi: " + raw);
                                }
                            }
                        }
                    }
                }
            }
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return toplam;
    }
    private boolean plakaGecerliMi(String plaka) {
        // Türk plakaları için: 34ABC123
        String turkRegex = "^(0[1-9]|[1-7][0-9]|8[0-1]) ?[A-Z]{1,3} ?\\d{2,4}$";

        // Alman plakaları için: B-AB 1234
        String almanRegex = "^[A-ZÄÖÜ]{1,3}-[A-Z]{1,2} \\d{1,4}$";

        // Taksi plakaları için: AC TX 711, AC TX 72, vb.
        String taksiRegex = "^[A-Z]{1,3} TX \\d{1,4}$";

        return plaka.matches(turkRegex) || plaka.matches(almanRegex) || plaka.matches(taksiRegex);
    }




private void initialize() {
    try {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
    } catch (Exception ignored) {}

    frame = new JFrame("🚗 Fahrtennachverfolgungssystem");
    frame.setSize(1400, 800);
    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    frame.setLocationRelativeTo(null);
    frame.setLayout(new BorderLayout(10, 10));

    Font anaFont = new Font("Segoe UI", Font.PLAIN, 14);
    JPanel ustPanel = new JPanel(new BorderLayout(10, 10));
    JPanel butonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 10));

    JButton btnSurusEkle = new JButton("Fahrt hinzufügen");
    JButton btnAddDriver = new JButton("Fahrer hinzufügen");
    JButton btnAddPlate = new JButton("Kennzeichen hinzufügen");
    JButton btnRapor = new JButton("Rapor erstellen");
    JButton btnFatura = new JButton("Fatura Oluştur");
    for (JButton b : new JButton[]{btnSurusEkle, btnAddDriver, btnAddPlate, btnRapor,btnFatura}) {
        b.setFont(anaFont);
        butonPanel.add(b);
    }


    formPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
    JPanel ustIcPanel = new JPanel(new BorderLayout());
    ustIcPanel.add(butonPanel, BorderLayout.WEST);

    JButton btnGuncelle = new JButton("Güncellemeleri Kaydet");
    btnGuncelle.setFont(anaFont);
    btnGuncelle.addActionListener(e -> {
        try {
            ExcelExporter.exportToExcel(tableModel, "surusler.xlsx");
            JOptionPane.showMessageDialog(frame, "Değişiklikler başarıyla Excel dosyasına kaydedildi.");
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(frame, "Kaydetme hatası: " + ex.getMessage());
        }
    });
    JButton btnFaturaExcel = new JButton("📥 Fatura Excel");
    btnFaturaExcel.setFont(anaFont);
    butonPanel.add(btnFaturaExcel);
    btnFaturaExcel.addActionListener(e -> setFormPanel(createFaturaExportPanel()));
    JPanel guncellePanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 10));
    guncellePanel.add(btnGuncelle);
    ustIcPanel.add(guncellePanel, BorderLayout.EAST);
    ustPanel.add(ustIcPanel, BorderLayout.NORTH);
    ustPanel.add(formPanel, BorderLayout.CENTER);
    frame.add(ustPanel, BorderLayout.NORTH);

    String[] columns = {
        "Sürücü", "Plaka", "Tarih", "Başlangıç KM", "Bitiş KM",
        "Bar", "RF", "Tanken", "Sonstig", "CC",
        "Umsatz", "Diff.-km", "Gef.-km", "Durs.-km", "Übergabe", "Vorschuss", "Sil"
    };

    tableModel = new DefaultTableModel(columns, 0) {
        @Override
        public boolean isCellEditable(int row, int column) {
            return column >= 0 && column <= 16;
        }
    };

    table = new JTable(tableModel) {
        @Override
        public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
            Component c = super.prepareRenderer(renderer, row, column);
            if (!isRowSelected(row)) {
                c.setBackground(row % 2 == 0 ? new Color(200, 220, 240) : Color.WHITE); // 🔵 Daha koyu mavi tonu
            } else {
                c.setBackground(new Color(173, 216, 230)); // Seçili satırın rengi
            }
            return c;
        }
    };
    table.setRowHeight(30);
    table.getColumn("Sil").setCellRenderer(new ButtonRenderer());
    table.getColumn("Sil").setCellEditor(new ButtonEditor(new JCheckBox(), table, tableModel));
    table.getColumn("Sil").setMaxWidth(50);

    JScrollPane scrollPane = new JScrollPane(table);
    frame.add(scrollPane);

    btnAddDriver.addActionListener(e -> setFormPanel(createDriverPanel()));
    btnAddPlate.addActionListener(e -> setFormPanel(createPlatePanel()));
    btnRapor.addActionListener(e -> setFormPanel(createReportPanel()));
    btnSurusEkle.addActionListener(e -> setFormPanel(createSurusPanel()));
    btnFatura.addActionListener(e -> setFormPanel(createFaturaPanel()));
    JTableHeader header = table.getTableHeader();
    header.addMouseListener(new MouseAdapter() {
        @Override
        public void mouseClicked(MouseEvent e) {
            int columnIndex = table.columnAtPoint(e.getPoint());
            String columnName = table.getColumnName(columnIndex);

            if (!(columnName.equals("Sürücü") || columnName.equals("Plaka") || columnName.equals("Tarih"))) return;

            Set<String> secenekler = new TreeSet<>();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                Object value = tableModel.getValueAt(i, columnIndex);
                if (value != null) secenekler.add(value.toString());
            }

            JPopupMenu popup = new JPopupMenu();

            JMenuItem tumunuGoster = new JMenuItem("Tümünü Göster");
            tumunuGoster.addActionListener(ev -> yenidenTabloyuYukle());
            popup.add(tumunuGoster);

            if (columnName.equals("Tarih")) {
                JPanel panel = new JPanel();
                JTextField txtTarih = new JTextField(10);
                JButton btnFiltrele = new JButton("Filtrele");

                btnFiltrele.addActionListener(ev -> {
                    String secili = txtTarih.getText().trim();
                    if (secili.isEmpty()) return;

                    for (int i = tableModel.getRowCount() - 1; i >= 0; i--) {
                        String cellValue = tableModel.getValueAt(i, columnIndex).toString();
                        if (!cellValue.equals(secili)) {
                            tableModel.removeRow(i);
                        }
                    }
                });

                panel.add(new JLabel("Tarih (gg.aa.yyyy):"));
                panel.add(txtTarih);
                panel.add(btnFiltrele);
                popup.add(panel);
            } else {
                for (String secenek : secenekler) {
                    JMenuItem item = new JMenuItem(secenek);
                    item.addActionListener(ev -> {
                        for (int i = tableModel.getRowCount() - 1; i >= 0; i--) {
                            String cellValue = tableModel.getValueAt(i, columnIndex).toString();
                            if (!cellValue.equals(secenek)) {
                                tableModel.removeRow(i);
                            }
                        }
                    });
                    popup.add(item);
                }
            }

            popup.show(header, e.getX(), header.getHeight());
        }
    });

    frame.setVisible(true);
    exceldenTabloyaOku("surusler.xlsx", tableModel);
}



private JPanel createFaturaPanel() {
    JPanel panel = new JPanel(new BorderLayout(10, 10));

    // 📤 Fatura Tablosu
    String[] columns = {"Tarih", "Başlık", "Vergi (%)", "Tutar (€)", "Vergi Türü"};
    DefaultTableModel model = new DefaultTableModel(columns, 0);
    JTable table = new JTable(model);
    JScrollPane scroll = new JScrollPane(table);

    // CSV'den veri yükle
    File f = new File("faturaverisi.csv");
    if (f.exists()) {
        try (Scanner sc = new Scanner(f)) {
            while (sc.hasNextLine()) {
                String[] data = sc.nextLine().split(",", -1);
                if (data.length == 5) model.addRow(data);
            }
        } catch (IOException ignored) {}
    }

    // 📋 Form Paneli
    JPanel formPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 10));
    formPanel.setBorder(BorderFactory.createTitledBorder("Yeni Fatura"));

    JSpinner dateSpinner = new JSpinner(new SpinnerDateModel());
    dateSpinner.setEditor(new JSpinner.DateEditor(dateSpinner, "dd.MM.yy"));
    JTextField txtBaslik = new JTextField(10);
    JComboBox<String> cbVergiOrani = new JComboBox<>(new String[]{"0", "7", "19"});
    JTextField txtTutar = new JTextField(8);
    JComboBox<String> cbVergiTuru = new JComboBox<>(new String[]{"RF", "ALT", "ASA", "NLR", "NET", "NORD"});
    cbVergiTuru.setEditable(true);
    JButton btnKaydet = new JButton("Faturayı Kaydet");

    formPanel.add(new JLabel("Tarih:"));
    formPanel.add(dateSpinner);
    formPanel.add(new JLabel("Başlık:"));
    formPanel.add(txtBaslik);
    formPanel.add(new JLabel("Vergi Oranı (%):"));
    formPanel.add(cbVergiOrani);
    formPanel.add(new JLabel("Tutar (€):"));
    formPanel.add(txtTutar);
    formPanel.add(new JLabel("Vergi Türü:"));
    formPanel.add(cbVergiTuru);
    formPanel.add(btnKaydet);

    // 🔍 FİLTRE PANELİ (sağ üstte)
    JPanel filtrePanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 10));
    filtrePanel.setBorder(BorderFactory.createTitledBorder("Filtrele"));

    JDateChooser dcBas = new JDateChooser(); dcBas.setDateFormatString("dd.MM.yy");
    JDateChooser dcBit = new JDateChooser(); dcBit.setDateFormatString("dd.MM.yy");
    JComboBox<String> cbFiltreVergiTuru = new JComboBox<>(new String[]{"Tümü", "RF", "ALT", "ASA", "NLR", "NET", "NORD"});
    JButton btnFiltrele = new JButton("Filtrele");
    

    filtrePanel.add(new JLabel("Başlangıç:"));
    filtrePanel.add(dcBas);
    filtrePanel.add(new JLabel("Bitiş:"));
    filtrePanel.add(dcBit);
    filtrePanel.add(new JLabel("Vergi Türü:"));
    filtrePanel.add(cbFiltreVergiTuru);
    filtrePanel.add(btnFiltrele);
 

    // 📊 ÖZET Panel
    JPanel ozetPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 10));
    ozetPanel.setBorder(BorderFactory.createTitledBorder("Filtre Sonucu Özeti"));

    JLabel lblToplamTutar = new JLabel("Toplam Tutar: €0.00");
    JLabel lblToplamVergi = new JLabel("Toplam Vergi: €0.00");
    JLabel lblKalanNet = new JLabel("Kalan Net: €0.00");

    ozetPanel.add(lblToplamTutar);
    ozetPanel.add(lblToplamVergi);
    ozetPanel.add(lblKalanNet);

    // 💾 KAYDET BUTONU
    btnKaydet.addActionListener(e -> {
        try {
            String tarih = new SimpleDateFormat("dd.MM.yy").format((Date) dateSpinner.getValue());
            String baslik = txtBaslik.getText().trim();
            String vergiOrani = cbVergiOrani.getSelectedItem().toString();
            String tutarStr = txtTutar.getText().trim().replace(",", ".");
            String vergiTuru = cbVergiTuru.getSelectedItem().toString();

            if (baslik.isEmpty() || tutarStr.isEmpty()) {
                JOptionPane.showMessageDialog(panel, "Lütfen tüm alanları doldurun.");
                return;
            }

            double tutar = Double.parseDouble(tutarStr);

            try (BufferedWriter writer = new BufferedWriter(new FileWriter("faturaverisi.csv", true))) {
                writer.write(String.join(",", tarih, baslik, vergiOrani, String.format(Locale.US, "%.2f", tutar), vergiTuru));
                writer.newLine();
            }

            model.addRow(new String[]{tarih, baslik, vergiOrani, String.format("%.2f", tutar), vergiTuru});
            txtBaslik.setText(""); txtTutar.setText("");

            JOptionPane.showMessageDialog(panel, "✅ Fatura kaydedildi.");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(panel, "Hata: " + ex.getMessage());
        }
    });

    // 📌 FİLTRELE BUTONU
    btnFiltrele.addActionListener(e -> {
        Date bas = dcBas.getDate();
        Date bit = dcBit.getDate();
        String seciliTur = cbFiltreVergiTuru.getSelectedItem().toString();

        if (bas == null || bit == null) {
            JOptionPane.showMessageDialog(panel, "Tarih aralığı seçin.");
            return;
        }

        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yy");
        List<String[]> filtreli = new ArrayList<>();

        try (Scanner sc = new Scanner(new File("faturaverisi.csv"))) {
            while (sc.hasNextLine()) {
                String[] data = sc.nextLine().split(",", -1);
                if (data.length != 5) continue;

                Date satirTarih = sdf.parse(data[0]);
                boolean tarihUygun = !satirTarih.before(bas) && !satirTarih.after(bit);
                boolean turUygun = seciliTur.equals("Tümü") || data[4].equalsIgnoreCase(seciliTur);

                if (tarihUygun && turUygun) {
                    filtreli.add(data);
                }
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(panel, "Filtreleme hatası: " + ex.getMessage());
            return;
        }

        // Tabloyu güncelle
        model.setRowCount(0);
        for (String[] satir : filtreli) model.addRow(satir);

        // Özet bilgileri hesapla
        double toplamTutar = 0, toplamVergi = 0;

        for (String[] satir : filtreli) {
            try {
                double tutar = Double.parseDouble(satir[3]);
                int oran = Integer.parseInt(satir[2]);
                toplamTutar += tutar;
                toplamVergi += (tutar * oran / 100.0);
            } catch (Exception ignored) {}
        }

        double kalan = toplamTutar - toplamVergi;
      

       

  

        lblToplamTutar.setText("Toplam Tutar: €" + String.format("%.2f", toplamTutar));
        lblToplamVergi.setText("Toplam Vergi: €" + String.format("%.2f", toplamVergi));
        lblKalanNet.setText("Kalan Net: €" + String.format("%.2f", kalan));
    });


    // 📌 Tümünü Göster Butonu
   

    // Ekrana yerleştir
    panel.add(formPanel, BorderLayout.NORTH);
    panel.add(scroll, BorderLayout.CENTER);

    JPanel ustSagPanel = new JPanel(new BorderLayout());
    ustSagPanel.add(filtrePanel, BorderLayout.NORTH);
    ustSagPanel.add(ozetPanel, BorderLayout.SOUTH);
    panel.add(ustSagPanel, BorderLayout.EAST);

    return panel;
}


   
   


    private JPanel createDriverPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

        // 🔼 Sürücü ekleme formu
        JPanel eklemeSatiri = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JLabel label = new JLabel("Sürücü İsmi: ");
        JTextField textField = new JTextField(20);
        JButton btnEkle = new JButton("Ekle");

        eklemeSatiri.add(label);
        eklemeSatiri.add(textField);
        eklemeSatiri.add(btnEkle);
        eklemeSatiri.setAlignmentX(Component.LEFT_ALIGNMENT); // Sola hizalama
        panel.add(eklemeSatiri);

        // 🔽 Butonun işlevi
        btnEkle.addActionListener(e -> {
            String adSoyad = textField.getText().trim();
            if (adSoyad.isEmpty()) {
                JOptionPane.showMessageDialog(frame, "Sürücü ismi boş olamaz.");
                return;
            }

            int onay = JOptionPane.showConfirmDialog(frame, adSoyad + " isimli sürücüyü eklemek istiyor musunuz?", "Onay", JOptionPane.YES_NO_OPTION);
            if (onay != JOptionPane.YES_OPTION) return;

            try (BufferedWriter writer = new BufferedWriter(new FileWriter("suruculer.csv", true))) {
                writer.write(adSoyad);
                writer.newLine();
                JOptionPane.showMessageDialog(frame, "✅ Sürücü eklendi: " + adSoyad);
                textField.setText("");

                // 🔄 Paneli yeniden yükle
                setFormPanel(createDriverPanel());

                // 🔁 Eğer sekmeli yapı varsa sürücü sekmesine geç
                // tabbedPane.setSelectedIndex(1);

            } catch (IOException ex) {
                JOptionPane.showMessageDialog(frame, "Hata: " + ex.getMessage());
            }
        });


        // 🔽 Kayıtlı sürücüleri göster
        panel.add(new JLabel("📋 Kayıtlı Sürücüler:"));

        JPanel surucuListesi = new JPanel();
        surucuListesi.setLayout(new BoxLayout(surucuListesi, BoxLayout.Y_AXIS));

        try (Scanner sc = new Scanner(new File("suruculer.csv"))) {
            while (sc.hasNextLine()) {
                String isim = sc.nextLine().trim();
                if (isim.isEmpty()) continue;

                JPanel satir = new JPanel(new FlowLayout(FlowLayout.LEFT));
                satir.add(new JLabel("👤 " + isim));

                JButton btnSil = new JButton("❌ Sil");
                btnSil.addActionListener(ev -> {
                    int onay = JOptionPane.showConfirmDialog(frame,
                            isim + " isimli sürücüyü silmek istiyor musunuz?",
                            "Silme Onayı", JOptionPane.YES_NO_OPTION);
                    if (onay == JOptionPane.YES_OPTION) {
                        try {
                            List<String> lines = new ArrayList<>();
                            Scanner s = new Scanner(new File("suruculer.csv"));
                            while (s.hasNextLine()) {
                                String line = s.nextLine();
                                if (!line.trim().equals(isim)) lines.add(line);
                            }
                            s.close();
                            java.nio.file.Files.write(new File("suruculer.csv").toPath(), lines);
                            JOptionPane.showMessageDialog(frame, "❌ Sürücü silindi: " + isim);
                            setFormPanel(createDriverPanel()); // Yeniden yükle
                        } catch (IOException ex) {
                            JOptionPane.showMessageDialog(frame, "Silme hatası: " + ex.getMessage());
                        }
                    }
                });

                satir.add(btnSil);
                satir.setAlignmentX(Component.LEFT_ALIGNMENT);
                surucuListesi.add(satir);
            }
        } catch (IOException ex) {
            surucuListesi.add(new JLabel("Yükleme hatası: " + ex.getMessage()));
        }

        JScrollPane scroll = new JScrollPane(surucuListesi);
        scroll.setPreferredSize(new Dimension(520, 200));
        scroll.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(scroll);
        
        return panel;
    }
    private JPanel createFaturaExportPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 15));
        panel.setBorder(BorderFactory.createTitledBorder("🧾 Fatura Excel Oluştur"));

        // Giriş alanları
        JLabel lblBas = new JLabel("Başlangıç Tarihi:");
        JDateChooser dcBas = new JDateChooser(); dcBas.setDateFormatString("dd.MM.yy");

        JLabel lblBit = new JLabel("Bitiş Tarihi:");
        JDateChooser dcBit = new JDateChooser(); dcBit.setDateFormatString("dd.MM.yy");

        JLabel lblTaxiNr = new JLabel("Taxi Nr:");
        JTextField tfTaxiNr = new JTextField(6);

        JLabel lblRechnung = new JLabel("Rechnungsnummer:");
        JTextField tfRechnung = new JTextField(10);

        JLabel lblLeistungszeit = new JLabel("Leistungszeitraum:");
        JTextField tfLeistung = new JTextField(10);

        JLabel lblIban = new JLabel("IBAN:");
        JTextField tfIban = new JTextField(20);

        JButton btnOlustur = new JButton("📤 Excel Fatura Oluştur");

        panel.add(lblBas); panel.add(dcBas);
        panel.add(lblBit); panel.add(dcBit);
        panel.add(lblTaxiNr); panel.add(tfTaxiNr);
        panel.add(lblRechnung); panel.add(tfRechnung);
        panel.add(lblLeistungszeit); panel.add(tfLeistung);
        panel.add(lblIban); panel.add(tfIban);
        panel.add(btnOlustur);

        btnOlustur.addActionListener(e -> {
            try {
                Date basDate = dcBas.getDate();
                Date bitDate = dcBit.getDate();

                if (basDate == null || bitDate == null) {
                    JOptionPane.showMessageDialog(panel, "Tarih aralığını seçiniz.");
                    return;
                }

                String taxiNr = tfTaxiNr.getText().trim();
                String rechnungNr = tfRechnung.getText().trim();
                String leistungszeit = tfLeistung.getText().trim();
                String iban = tfIban.getText().trim();
                String tarih = new SimpleDateFormat("dd.MM.yyyy").format(new Date());

                if (taxiNr.isEmpty() || rechnungNr.isEmpty() || leistungszeit.isEmpty() || iban.isEmpty()) {
                    JOptionPane.showMessageDialog(panel, "Tüm alanları doldurun.");
                    return;
                }

                // csv'den oku ve vergi türlerine göre hesapla
                double net7 = 0, tax7 = 0;
                double net19 = 0, tax19 = 0;
                double altNet19 = 0, altTax19 = 0;

                SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yy");

                try (Scanner sc = new Scanner(new File("faturaverisi.csv"))) {
                    while (sc.hasNextLine()) {
                        String[] data = sc.nextLine().split(",", -1);
                        if (data.length != 5) continue;

                        Date tarihSatir = sdf.parse(data[0]);
                        if (tarihSatir.before(basDate) || tarihSatir.after(bitDate)) continue;

                        int oran = Integer.parseInt(data[2]);
                        double tutar = Double.parseDouble(data[3]);
                        String tur = data[4];

                        double vergi = (tutar * oran) / (100 + oran);  // ters KDV
                        double net = tutar - vergi;

                        if (oran == 7) {
                            net7 += net;
                            tax7 += vergi;
                        } else if (oran == 19) {
                            if (tur.equals("ALT") || tur.equals("ASA") || tur.equals("RED")) {
                                altNet19 += net;
                                altTax19 += vergi;
                            } else {
                                net19 += net;
                                tax19 += vergi;
                            }
                        }
                    }
                }

                ExcelFaturaOlustur.olustur(
                        taxiNr, rechnungNr, leistungszeit, tarih, iban,
                        net7, tax7, net19, tax19, altNet19, altTax19
                );

                JOptionPane.showMessageDialog(panel, "✅ Excel faturası oluşturuldu → faturalar.xlsx");

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(panel, "Hata: " + ex.getMessage());
            }
        });

        return panel;
    }






    //plaka ekleme paneli 
    
 private JPanel createPlatePanel() {
    JPanel panel = new JPanel();
    panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
    panel.setAlignmentX(Component.LEFT_ALIGNMENT);

    // 🔼 PLAKA EKLEME FORMU
    JPanel eklemeSatiri = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
    JTextField txtPlaka = new JTextField(8);
    JTextField txtKm = new JTextField(6);
    JComboBox<String> aySecim = new JComboBox<>(new String[]{
    	    "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
    	    "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
    	});
    JButton btnEkle = new JButton("Ekle");

    eklemeSatiri.add(new JLabel("Yeni Plaka:"));
    eklemeSatiri.add(txtPlaka);
    eklemeSatiri.add(new JLabel("Başlangıç KM:"));
    eklemeSatiri.add(txtKm);
    eklemeSatiri.add(new JLabel("Muayene Ayı:"));
    eklemeSatiri.add(aySecim);
    eklemeSatiri.add(btnEkle);
    eklemeSatiri.setAlignmentX(Component.LEFT_ALIGNMENT);
    panel.add(eklemeSatiri);

    // 🔽 Ekleme butonu işlevi
    btnEkle.addActionListener(e -> {
        String plaka = txtPlaka.getText().trim().toUpperCase();
        String kmText = txtKm.getText().trim();
        int bakim = 0;

        if (!plakaGecerliMi(plaka)) {
            JOptionPane.showMessageDialog(frame,
                    "Geçersiz plaka formatı!\n\nGeçerli Türk veya Alman plakaları giriniz.\n(örn. 34ABC123 veya B-AB 1234)",
                    "Plaka Hatası", JOptionPane.ERROR_MESSAGE);
            return;
        }

        try {
            int kmDeger = Integer.parseInt(kmText);
            int secilenAy = aySecim.getSelectedIndex() + 1;

            try (BufferedWriter writer = new BufferedWriter(new FileWriter("plakalar.csv", true))) {
                writer.write(plaka + "," + kmDeger + "," + kmDeger + "," + bakim + "," + secilenAy);
                writer.newLine();
                JOptionPane.showMessageDialog(frame, "✅ Plaka eklendi: " + plaka + " (Muayene Ayı: " + secilenAy + ")");
                setFormPanel(createPlatePanel());
            }
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(frame, "Başlangıç KM sayısal olmalı.");
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(frame, "Dosya hatası: " + ex.getMessage());
        }
    });

    panel.add(new JLabel("📋 Kayıtlı Plakalar:"));

    JPanel plakaListesi = new JPanel();
    plakaListesi.setLayout(new BoxLayout(plakaListesi, BoxLayout.Y_AXIS));
    plakaListesi.setAlignmentX(Component.LEFT_ALIGNMENT);

    try (Scanner sc = new Scanner(new File("plakalar.csv"))) {
        while (sc.hasNextLine()) {
            String[] parts = sc.nextLine().split(",");
            if (parts.length >= 5) {
                String plaka = parts[0];
                int basKm = Integer.parseInt(parts[1]);
                int sonKm = Integer.parseInt(parts[2]);
                int bakimDurum = Integer.parseInt(parts[3]);
                int muayeneAyi = Integer.parseInt(parts[4]);

                int fark = sonKm - basKm;
                int kalanKm = 15000 - (sonKm % 15000);
                if (kalanKm == 0) kalanKm = 15000;

                int toplamGelir = hesaplaToplamGelir(plaka);

                JPanel satir = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
                satir.setAlignmentX(Component.LEFT_ALIGNMENT);
                satir.add(new JLabel("🚘 " + plaka));
                satir.add(new JLabel(" | Son KM: " + sonKm));
                satir.add(new JLabel(" | Bakım: " + (bakimDurum == 1 ? "✔️" : "❌")));
                satir.add(new JLabel(" | Bakıma kalan: " + kalanKm + " km"));
                satir.add(new JLabel(" | Muayene Ayı: " + muayeneAyi));
                satir.add(new JLabel(" | Gelir: €" + toplamGelir));

                JButton btnBakimYap = new JButton("🛠️ Bakım Yap");
                JButton btnBakimsizYap = new JButton("❌ Bakımsız Yap");
                JButton btnSil = new JButton("Sil");

                btnBakimYap.addActionListener(ev -> {
                    updateBakimDurumu(plaka, 1);
                    JOptionPane.showMessageDialog(frame, plaka + " için bakım yapıldı.");
                    setFormPanel(createPlatePanel());
                });

                btnBakimsizYap.addActionListener(ev -> {
                    updateBakimDurumu(plaka, 0);
                    JOptionPane.showMessageDialog(frame, plaka + " bakımsız yapıldı.");
                    setFormPanel(createPlatePanel());
                });

                btnSil.addActionListener(ev -> {
                    int secim = JOptionPane.showConfirmDialog(
                            frame,
                            plaka + " plakalı aracı silmek istiyor musunuz?",
                            "Silme Onayı",
                            JOptionPane.YES_NO_OPTION
                    );

                    if (secim == JOptionPane.YES_OPTION) {
                        try {
                            List<String> lines = new ArrayList<>();
                            try (Scanner reader = new Scanner(new File("plakalar.csv"))) {
                                while (reader.hasNextLine()) {
                                    String satirData = reader.nextLine();
                                    if (!satirData.startsWith(plaka + ",")) {
                                        lines.add(satirData);
                                    }
                                }
                            }
                            java.nio.file.Files.write(new File("plakalar.csv").toPath(), lines);
                            JOptionPane.showMessageDialog(frame, "🚫 Plaka silindi: " + plaka);
                            setFormPanel(createPlatePanel());
                        } catch (IOException ex) {
                            JOptionPane.showMessageDialog(frame, "Silme hatası: " + ex.getMessage());
                        }
                    }
                });

                satir.add(btnBakimYap);
                satir.add(btnBakimsizYap);
                satir.add(btnSil);
                plakaListesi.add(satir);
            }
        }
    } catch (Exception ex) {
        plakaListesi.add(new JLabel("Yükleme hatası: " + ex.getMessage()));
    }

    JScrollPane scroll = new JScrollPane(plakaListesi);
    scroll.setPreferredSize(new Dimension(940, 160));
    scroll.setMaximumSize(new Dimension(940, 160));
    scroll.setAlignmentX(Component.LEFT_ALIGNMENT);
    scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

    Box scrollWrapper = Box.createVerticalBox();
    scrollWrapper.setAlignmentX(Component.LEFT_ALIGNMENT);
    scrollWrapper.add(scroll);

    panel.add(scrollWrapper);

    return panel;
}







/*
 * 
 * 
 * 
 * 
 * 
 * SURUŞ PANELİİİİİİİİİİİİİİİİ
 * 
 * 
 * 
 * 
 * */
  private JPanel createSurusPanel() {
    JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
    JComboBox<String> cbFahrer = new JComboBox<>();
    JComboBox<String> cbPlaka = new JComboBox<>();

    // CSV'den sürücü ve plaka yükle
    try {
        Scanner scanner = new Scanner(new File("suruculer.csv"));
        while (scanner.hasNextLine()) cbFahrer.addItem(scanner.nextLine().trim());
        scanner.close();

        scanner = new Scanner(new File("plakalar.csv"));
        while (scanner.hasNextLine()) {
            String[] parts = scanner.nextLine().split(",");
            if (parts.length >= 5) cbPlaka.addItem(parts[0]);
        }
        scanner.close();
    } catch (Exception e) {
        JOptionPane.showMessageDialog(null, "Dosya okuma hatası: " + e.getMessage());
    }

    // Giriş alanları
    JDateChooser dateChooser = new JDateChooser();
    dateChooser.setDateFormatString("dd.MM.yyyy");

    JTextField tfStartKm = new JTextField(5);
    JTextField tfEndKm = new JTextField(5);
    JTextField tfBar = new JTextField(5);
    JTextField tfRf = new JTextField(5);
    JTextField tfCc = new JTextField(5);
    JTextField tfTanken = new JTextField(5);
    JTextField tfSonstig = new JTextField(5);
    JButton btnEkle = new JButton("Ekle");

    btnEkle.addActionListener(e -> {
        try {
            String surucu = cbFahrer.getSelectedItem().toString();
            String plaka = cbPlaka.getSelectedItem().toString();
            Date seciliTarih = dateChooser.getDate();

            if (seciliTarih == null) {
                JOptionPane.showMessageDialog(null, "Lütfen geçerli bir tarih seçin.");
                return;
            }

            String tarih = new SimpleDateFormat("dd.MM.yyyy").format(seciliTarih);
            int girilenAy = Integer.parseInt(new SimpleDateFormat("MM").format(seciliTarih));

            int basKm = Integer.parseInt(tfStartKm.getText().trim());
            int bitKm = Integer.parseInt(tfEndKm.getText().trim());
            double bar = Double.parseDouble(tfBar.getText().trim().replace(",", "."));
            double rf = Double.parseDouble(tfRf.getText().trim().replace(",", "."));
            double cc = Double.parseDouble(tfCc.getText().trim().replace(",", "."));
            double tanken = Double.parseDouble(tfTanken.getText().trim().replace(",", "."));
            double sonstig = Double.parseDouble(tfSonstig.getText().trim().replace(",", "."));

            if (bitKm < basKm)
                throw new Exception("Bitiş kilometresi, başlangıçtan küçük olamaz!");

            // plakalar.csv'den kontrol için değerleri al
            int sonKmPlakada = 0;
            int bakimDurumu = 0;
            int eskiBasKm = 0;
            int muayeneAyi = -1;

            try (Scanner sc = new Scanner(new File("plakalar.csv"))) {
                while (sc.hasNextLine()) {
                    String[] parts = sc.nextLine().split(",");
                    if (parts[0].equals(plaka) && parts.length >= 5) {
                        eskiBasKm = Integer.parseInt(parts[1]);
                        sonKmPlakada = Integer.parseInt(parts[2]);
                        bakimDurumu = Integer.parseInt(parts[3]);
                        muayeneAyi = Integer.parseInt(parts[4]);
                        break;
                    }
                }
            }

            // Muayene ayı kontrolü
            if (girilenAy == muayeneAyi) {
                JOptionPane.showMessageDialog(null,
                    "🚨 ACİLEN MUAYENEYE GİTMELİSİNİZ!\n\nPlaka: " + plaka + "\nBu ay araç muayenesi var!",
                    "Muayene Uyarısı", JOptionPane.WARNING_MESSAGE);
            }

            if (bakimDurumu == 0) {
                JOptionPane.showMessageDialog(null,
                    "⚠️ Bu araç şu anda BAKIMSIZ durumdadır!\nPlaka: " + plaka,
                    "Bakım Uyarısı", JOptionPane.WARNING_MESSAGE);
            }

            if (basKm < sonKmPlakada) {
                JOptionPane.showMessageDialog(null,
                    "Başlangıç KM, kayıtlı son KM'den küçük olamaz!\n\n" +
                    "Plaka: " + plaka + "\nSon kayıtlı KM: " + sonKmPlakada,
                    "KM Hatası", JOptionPane.WARNING_MESSAGE);
                return;
            }

            int gefKm = bitKm - basKm;
            int diffKm = gefKm;
            double umsatz = bar + rf + cc;
            double dursKm = (umsatz / gefKm);
            double vorschuss = umsatz * 0.4;
            double teslim = umsatz - vorschuss - sonstig - tanken;

            Object[] row = {
                surucu, plaka, tarih, basKm, bitKm,
                bar, rf, tanken, sonstig, cc,
                String.format("%.2f", umsatz),
                diffKm, gefKm,
                String.format("%.2f", dursKm),
                String.format("%.2f", teslim),
                String.format("%.2f", vorschuss),
                "✖"
            };
            tableModel.insertRow(0, row);

            // plakalar.csv güncelle
            List<String> lines = new ArrayList<>();
            java.nio.file.Path path = java.nio.file.Paths.get("plakalar.csv");
            java.nio.file.Files.lines(path).forEach(line -> {
                String[] parts = line.split(",");
                if (parts.length >= 5 && parts[0].equals(plaka)) {
                    int fark = bitKm - Integer.parseInt(parts[1]);
                    int yeniBakim = (fark >= 15000) ? 0 : Integer.parseInt(parts[3]);
                    int yeniBas = (fark >= 15000) ? bitKm : Integer.parseInt(parts[1]);
                    lines.add(plaka + "," + yeniBas + "," + bitKm + "," + yeniBakim + "," + parts[4]);
                } else {
                    lines.add(line);
                }
            });
            java.nio.file.Files.write(path, lines);

            ExcelExporter.exportToExcel(tableModel, "surusler.xlsx");

            JOptionPane.showMessageDialog(null, "Sürüş eklendi ve Excel dosyası güncellendi.");

            // Temizle
            tfStartKm.setText("");
            tfEndKm.setText("");
            tfBar.setText("");
            tfRf.setText("");
            tfCc.setText("");
            tfTanken.setText("");
            tfSonstig.setText("");
            dateChooser.setDate(null);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Hata: " + ex.getMessage());
        }
    });

    panel.add(new JLabel("Fahrer:")); panel.add(cbFahrer);
    panel.add(new JLabel("Kennzeichen:")); panel.add(cbPlaka);
    panel.add(new JLabel("Datum:")); panel.add(dateChooser);
    panel.add(new JLabel("Start Km:")); panel.add(tfStartKm);
    panel.add(new JLabel("Ende Km:")); panel.add(tfEndKm);
    panel.add(new JLabel("BAR (€):")); panel.add(tfBar);
    panel.add(new JLabel("RF (€):")); panel.add(tfRf);
    panel.add(new JLabel("CC:")); panel.add(tfCc);
    panel.add(new JLabel("TANKEN (€):")); panel.add(tfTanken);
    panel.add(new JLabel("SONSTIG (€):")); panel.add(tfSonstig);
    panel.add(btnEkle);

    return panel;
}




    private void yenidenTabloyuYukle() {
        exceldenTabloyaOku("surusler.xlsx", tableModel);
    }

    // setFormPanel metodu güncellendi
  

    // createReportPanel metodu güncellendi
 // createReportPanel metodu güncel hali
 // createReportPanel metodu güncel hali
  private JPanel createReportPanel() {
    JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));

    JLabel lblBas = new JLabel("Başlangıç Tarihi: ");
    JDateChooser baslangicChooser = new JDateChooser();
    baslangicChooser.setDateFormatString("dd.MM.yyyy");

    JLabel lblBit = new JLabel("Bitiş Tarihi: ");
    JDateChooser bitisChooser = new JDateChooser();
    bitisChooser.setDateFormatString("dd.MM.yyyy");

    JLabel lblPlaka = new JLabel("Plaka:");
    JComboBox<String> cbPlaka = new JComboBox<>();
    cbPlaka.addItem("Tümü");

    JLabel lblSurucu = new JLabel("Sürücü:");
    JComboBox<String> cbSurucu = new JComboBox<>();
    cbSurucu.addItem("Tümü");

    // Plaka ve sürücüleri yükle
    try (Scanner sc = new Scanner(new File("plakalar.csv"))) {
        while (sc.hasNextLine()) {
            String[] parts = sc.nextLine().split(",");
            if (parts.length > 0) cbPlaka.addItem(parts[0]);
        }
    } catch (Exception e) {
        JOptionPane.showMessageDialog(frame, "Plaka yüklenemedi: " + e.getMessage());
    }

    try (Scanner sc = new Scanner(new File("suruculer.csv"))) {
        while (sc.hasNextLine()) {
            cbSurucu.addItem(sc.nextLine().trim());
        }
    } catch (Exception e) {
        JOptionPane.showMessageDialog(frame, "Sürücü yüklenemedi: " + e.getMessage());
    }

    JButton btnFiltrele = new JButton("Filtrele");

    // Renkli özet paneli
    JPanel summaryPanel = new JPanel(new GridLayout(0, 4, 10, 6));
    summaryPanel.setBorder(BorderFactory.createTitledBorder("⚡ Filtre Sonuç Özeti"));
    summaryPanel.setBackground(new Color(245, 245, 245));
    Font ozetFont = new Font("Segoe UI", Font.BOLD, 14);

    JLabel lblToplamUmsatz = new JLabel("Umsatz: €0.00");
    lblToplamUmsatz.setForeground(new Color(33, 150, 243)); // Mavi

    JLabel lblToplamTanken = new JLabel("Tanken: €0.00");
    lblToplamTanken.setForeground(new Color(255, 87, 34)); // Turuncu kırmızı

    JLabel lblToplamSonstig = new JLabel("Sonstig: €0.00");
    lblToplamSonstig.setForeground(new Color(244, 67, 54)); // Kırmızı

    JLabel lblToplamTeslim = new JLabel("Übergabe: €0.00");
    lblToplamTeslim.setForeground(new Color(76, 175, 80)); // Yeşil

    JLabel lblToplamVorschuss = new JLabel("Vorschuss: €0.00");
    lblToplamVorschuss.setForeground(new Color(156, 39, 176)); // Mor

    JLabel lblToplamBar = new JLabel("Bar: €0.00");
    lblToplamBar.setForeground(new Color(255, 152, 0)); // Turuncu

    JLabel lblToplamRF = new JLabel("RF: €0.00");
    lblToplamRF.setForeground(new Color(0, 188, 212)); // Camgöbeği

    JLabel lblToplamCC = new JLabel("CC: €0.00");
    lblToplamCC.setForeground(new Color(121, 85, 72)); // Kahverengi

    JLabel lblToplamDursKm = new JLabel("Durs.-km: 0.00");
    lblToplamDursKm.setForeground(new Color(96, 125, 139)); // Gri-mavi

    for (JLabel lbl : new JLabel[]{
            lblToplamBar, lblToplamRF, lblToplamCC, lblToplamDursKm,
            lblToplamUmsatz, lblToplamTanken, lblToplamSonstig,
            lblToplamTeslim, lblToplamVorschuss
    }) {
        lbl.setFont(ozetFont);
        lbl.setBorder(new EmptyBorder(4, 8, 4, 8));
        summaryPanel.add(lbl);
    }

    // Filtreleme butonu aksiyonu
    btnFiltrele.addActionListener(e -> {
        yenidenTabloyuYukle();

        Date basDate = baslangicChooser.getDate();
        Date bitDate = bitisChooser.getDate();
        String secilenPlaka = cbPlaka.getSelectedItem().toString();
        String secilenSurucu = cbSurucu.getSelectedItem().toString();

        if (basDate == null || bitDate == null) {
            JOptionPane.showMessageDialog(frame, "Lütfen geçerli bir tarih aralığı seçiniz.");
            return;
        }

        SimpleDateFormat df = new SimpleDateFormat("dd.MM.yyyy");

        for (int i = tableModel.getRowCount() - 1; i >= 0; i--) {
            try {
                Date rowDate = df.parse(tableModel.getValueAt(i, 2).toString());
                String plaka = tableModel.getValueAt(i, 1).toString();
                String surucu = tableModel.getValueAt(i, 0).toString();

                boolean tarihUygun = !rowDate.before(basDate) && !rowDate.after(bitDate);
                boolean plakaUygun = secilenPlaka.equals("Tümü") || plaka.equals(secilenPlaka);
                boolean surucuUygun = secilenSurucu.equals("Tümü") || surucu.equals(secilenSurucu);

                if (!(tarihUygun && plakaUygun && surucuUygun)) {
                    tableModel.removeRow(i);
                }
            } catch (Exception ex) {
                tableModel.removeRow(i);
            }
        }

        // Hesaplamalar
        double toplamUmsatz = 0, toplamTanken = 0, toplamSonstig = 0, toplamTeslim = 0, toplamVorschuss = 0;
        double toplamBar = 0, toplamRF = 0, toplamCC = 0, toplamDursKm = 0;
        int dursKmSayac = 0;

        for (int i = 0; i < tableModel.getRowCount(); i++) {
            try {
                toplamBar += Double.parseDouble(tableModel.getValueAt(i, 5).toString().replace(",", "."));
                toplamRF += Double.parseDouble(tableModel.getValueAt(i, 6).toString().replace(",", "."));
                toplamTanken += Double.parseDouble(tableModel.getValueAt(i, 7).toString().replace(",", "."));
                toplamSonstig += Double.parseDouble(tableModel.getValueAt(i, 8).toString().replace(",", "."));
                toplamCC += Double.parseDouble(tableModel.getValueAt(i, 9).toString().replace(",", "."));
                toplamUmsatz += Double.parseDouble(tableModel.getValueAt(i, 10).toString().replace(",", "."));
                toplamTeslim += Double.parseDouble(tableModel.getValueAt(i, 14).toString().replace(",", "."));
                toplamVorschuss += Double.parseDouble(tableModel.getValueAt(i, 15).toString().replace(",", "."));

                double dursKm = Double.parseDouble(tableModel.getValueAt(i, 13).toString().replace(",", "."));
                toplamDursKm += dursKm;
                dursKmSayac++;
            } catch (Exception ignored) {}
        }

        double ortDursKm = dursKmSayac > 0 ? toplamDursKm / dursKmSayac : 0;

        // Etiketleri güncelle
        lblToplamBar.setText("Bar: €" + String.format("%.2f", toplamBar));
        lblToplamRF.setText("RF: €" + String.format("%.2f", toplamRF));
        lblToplamCC.setText("CC: €" + String.format("%.2f", toplamCC));
        lblToplamUmsatz.setText("Umsatz: €" + String.format("%.2f", toplamUmsatz));
        lblToplamTanken.setText("Tanken: €" + String.format("%.2f", toplamTanken));
        lblToplamSonstig.setText("Sonstig: €" + String.format("%.2f", toplamSonstig));
        lblToplamTeslim.setText("Übergabe: €" + String.format("%.2f", toplamTeslim));
        lblToplamVorschuss.setText("Vorschuss: €" + String.format("%.2f", toplamVorschuss));
        lblToplamDursKm.setText("Durs.-km: " + String.format("%.2f", ortDursKm));
    });

    // Panel bileşenleri
    panel.add(lblBas); panel.add(baslangicChooser);
    panel.add(lblBit); panel.add(bitisChooser);
    panel.add(lblPlaka); panel.add(cbPlaka);
    panel.add(lblSurucu); panel.add(cbSurucu);
    panel.add(btnFiltrele);
    panel.add(summaryPanel);

    return panel;
}





    
    public static void exceldenTabloyaOku(String path, DefaultTableModel model) {
        try (FileInputStream fis = new FileInputStream(new File(path));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            model.setRowCount(0); // Tabloyu temizle

            // Mevcut sürücüleri oku
            List<String> mevcutSuruculer = new ArrayList<>();
            File surucuDosyasi = new File("suruculer.csv");
            if (surucuDosyasi.exists()) {
                try (Scanner sc = new Scanner(surucuDosyasi)) {
                    while (sc.hasNextLine()) {
                        mevcutSuruculer.add(sc.nextLine().trim().toLowerCase());
                    }
                }
            }

            // Yeni sürücüleri dosyaya eklemek için yazıcı
            BufferedWriter writer = new BufferedWriter(new FileWriter("suruculer.csv", true));

            java.text.SimpleDateFormat hedefFormat = new java.text.SimpleDateFormat("dd.MM.yyyy");
            String[] tarihFormatlari = {
                "d.M.yy", "d.M.yyyy", "dd.MM.yy", "dd.MM.yyyy",
                "d/M/yy", "d/M/yyyy", "dd/MM/yy", "dd/MM/yyyy",
                "yyyy-MM-dd", "yy.MM.dd"
            };

            for (int i = sheet.getLastRowNum(); i >= 1; i--) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Object[] rowData = new Object[17];
                for (int j = 0; j < rowData.length; j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null) {
                        rowData[j] = (j == 0) ? "" : 0;
                    } else {
                        if (j == 2) { // 📅 Tarih sütunu
                            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                                java.util.Date date = cell.getDateCellValue();
                                rowData[j] = hedefFormat.format(date);
                            } else if (cell.getCellType() == CellType.STRING) {
                                String rawDate = cell.getStringCellValue().trim();
                                boolean parsed = false;
                                for (String format : tarihFormatlari) {
                                    try {
                                        java.util.Date parsedDate = new java.text.SimpleDateFormat(format).parse(rawDate);

                                        // 🚫 2026 yılı kontrolü
                                        if (parsedDate.getYear() + 1900 >= 2026) {
                                        
                                            workbook.close();
                                            fis.close();
                                            return;
                                        }

                                        rowData[j] = hedefFormat.format(parsedDate);
                                        parsed = true;
                                        break;
                                    } catch (Exception ignored) {}
                                }
                                if (!parsed) rowData[j] = rawDate;
                            } else {
                                rowData[j] = "";
                            }
                        } else {
                            switch (cell.getCellType()) {
                                case NUMERIC -> rowData[j] = cell.getNumericCellValue();
                                case STRING -> rowData[j] = cell.getStringCellValue();
                                default -> rowData[j] = (j == 0) ? "" : 0;
                            }
                        }
                    }
                }

                // Sürücü kontrolü ve eklemesi
                String surucu = rowData[0].toString().trim();
                if (surucu.isEmpty() || surucu.equals("0")) continue;
                if (!surucu.isEmpty()) {
                    String normalized = surucu.toLowerCase();
                    if (!mevcutSuruculer.contains(normalized)) {
                        writer.write(surucu);
                        writer.newLine();
                        mevcutSuruculer.add(normalized);
                    }
                }

                model.insertRow(0, rowData);
            }

            writer.close();

        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Excel'den veri okuma hatası: " + e.getMessage());
        }
    }





    public static void main(String[] args) {
        SwingUtilities.invokeLater(anamenü::new);
    }
}
