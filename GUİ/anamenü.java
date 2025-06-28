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
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
//Ekstra (JTextComponent, EventListener vs.)
import javax.swing.text.JTextComponent;



//AWT
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
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

//Veri yapıları
import java.util.ArrayList;
import java.util.List;
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
        	ExcelExporter.exportToExcel(tableModel, seciliExcelYolu);
        	
      
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
                if (parts[0].equals(plaka)) {
                	parts[1]=parts[2];
                    lines.add(parts[0] + "," + parts[1] + "," + parts[2] + "," + yeniDurum);
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
        try (FileInputStream fis = new FileInputStream(seciliExcelYolu)) {
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

        for (JButton b : new JButton[]{btnSurusEkle, btnAddDriver, btnAddPlate, btnRapor}) {
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
                ExcelExporter.exportToExcel(tableModel, seciliExcelYolu);
                JOptionPane.showMessageDialog(frame, "Değişiklikler başarıyla Excel dosyasına kaydedildi.");
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(frame, "Kaydetme hatası: " + ex.getMessage());
            }
        });

        JButton btnExcelSec = new JButton("Excel Dosyası Seç");
        btnExcelSec.setFont(anaFont);
        btnExcelSec.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int secim = fileChooser.showOpenDialog(frame);
            if (secim == JFileChooser.APPROVE_OPTION) {
                File secilenDosya = fileChooser.getSelectedFile();
                seciliExcelYolu = secilenDosya.getAbsolutePath();
                exceldenTabloyaOku(seciliExcelYolu, tableModel);
                JOptionPane.showMessageDialog(frame, "Excel dosyası yüklendi:\n" + seciliExcelYolu);
            }
        });

        JPanel guncellePanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 10));
        guncellePanel.add(btnGuncelle);
        guncellePanel.add(btnExcelSec);

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

        table = new JTable(tableModel);
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
                }
 else {
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
        exceldenTabloyaOku(seciliExcelYolu, tableModel);
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






    //plaka ekleme paneli 
    
    private JPanel createPlatePanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setAlignmentX(Component.LEFT_ALIGNMENT);

        // 🔼 PLAKA EKLEME FORMU
        JPanel eklemeSatiri = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        JTextField txtPlaka = new JTextField(8);
        JTextField txtKm = new JTextField(6);
        JButton btnEkle = new JButton("Ekle");

        eklemeSatiri.add(new JLabel("Yeni Plaka:"));
        eklemeSatiri.add(txtPlaka);
        eklemeSatiri.add(new JLabel("Başlangıç KM:"));
        eklemeSatiri.add(txtKm);
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
                try (BufferedWriter writer = new BufferedWriter(new FileWriter("plakalar.csv", true))) {
                    writer.write(plaka + "," + kmDeger + "," + kmDeger + "," + bakim);
                    writer.newLine();
                    JOptionPane.showMessageDialog(frame, "✅ Plaka eklendi: " + plaka + " (bakımsız)");
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
                if (parts.length >= 4) {
                    String plaka = parts[0];
                    int basKm = Integer.parseInt(parts[1]);
                    int sonKm = Integer.parseInt(parts[2]);
                    int bakimDurum = Integer.parseInt(parts[3]);

                    int kalanKm = 15000 - (sonKm - basKm);
                    int toplamGelir = hesaplaToplamGelir(plaka);

                    JPanel satir = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
                    satir.setAlignmentX(Component.LEFT_ALIGNMENT);
                    satir.add(new JLabel("🚘 " + plaka));
                    satir.add(new JLabel(" | Son KM: " + sonKm));
                    satir.add(new JLabel(" | Bakım: " + (bakimDurum == 1 ? "✔️" : "❌")));
                    satir.add(new JLabel(" | Bakıma kalan: " + kalanKm + " km"));
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
                if (parts.length > 0) cbPlaka.addItem(parts[0]);
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

        // Ekle butonu
        btnEkle.addActionListener(e -> {
            try {
                String surucu = cbFahrer.getSelectedItem().toString();
                String plaka = cbPlaka.getSelectedItem().toString();
                String tarih = new java.text.SimpleDateFormat("dd.MM.yyyy").format(dateChooser.getDate());

                int basKm = Integer.parseInt(tfStartKm.getText().trim());
                int bitKm = Integer.parseInt(tfEndKm.getText().trim());
                double bar = Double.parseDouble(tfBar.getText().trim());
                double rf = Double.parseDouble(tfRf.getText().trim());
                double cc = Double.parseDouble(tfCc.getText().trim());
                double tanken = Double.parseDouble(tfTanken.getText().trim());
                double sonstig = Double.parseDouble(tfSonstig.getText().trim());

                if (bitKm < basKm)
                    throw new Exception("Ende-Kilometer darf nicht kleiner als Start-Kilometer sein.");

                // Son KM ve bakım kontrolü
                int sonKmPlakada = 0;
                int bakimDurumu = 0;
                try (Scanner sc = new Scanner(new File("plakalar.csv"))) {
                    while (sc.hasNextLine()) {
                        String[] parts = sc.nextLine().split(",");
                        if (parts[0].equals(plaka) && parts.length >= 4) {
                            sonKmPlakada = Integer.parseInt(parts[2]);
                            bakimDurumu = Integer.parseInt(parts[3]);
                            break;
                        }
                    }
                }

                if (bakimDurumu == 0) {
                    JOptionPane.showMessageDialog(null,
                        "⚠️ Dikkat: Bu araç şu anda **bakımsız** durumdadır!\n\n" +
                        "Lütfen en kısa sürede bakım yapınız.\n" +
                        "Plaka: " + plaka,
                        "Bakım Uyarısı", JOptionPane.WARNING_MESSAGE);
                }

                if (basKm < sonKmPlakada) {
                    JOptionPane.showMessageDialog(null,
                        "❗ Başlangıç kilometresi, bu plakaya ait en son kaydedilen kilometreden küçük olamaz!\n\n" +
                        "📍 Plaka: " + plaka + "\n" +
                        "🔢 Son kayıtlı kilometre: " + sonKmPlakada + " km\n" +
                        "🛑 Girilen başlangıç kilometre: " + basKm + " km\n\n" +
                        "Lütfen doğru ve güncel bir değer giriniz.",
                        "KM Uyuşmazlığı", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                // Hesaplamalar
                int gefKm = bitKm - basKm;
                int diffKm = gefKm;
                double umsatz = bar + rf  + cc ;
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
                    if (parts[0].equals(plaka)) {
                        int eskiBas = Integer.parseInt(parts[1]);
                        int bakim = Integer.parseInt(parts[3]);
                        int yeniBakim;
                        int yeniBas;

                        int fark = bitKm - eskiBas;

                        if (fark >= 15000) {
                            yeniBakim = 0;      // bakım zamanı geldi, artık bakımsız
                            yeniBas = bitKm;    // başlangıç km, yeni bitiş km'ye çekilir
                        } else {
                            yeniBakim = bakim;  // önceki bakım durumu korunur
                            yeniBas = eskiBas;  // başlangıç km olduğu gibi kalır
                        }
                        lines.add(plaka + "," + eskiBas+ "," + bitKm + "," + yeniBakim);
                    } else {
                        lines.add(line);
                    }
                });
                java.nio.file.Files.write(path, lines);

                // Excel güncelle
                ExcelExporter.exportToExcel(tableModel, seciliExcelYolu);

                JOptionPane.showMessageDialog(null, "Sürüş eklendi ve Excel dosyası güncellendi.");

                // Alanları temizle
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

        // Arayüz
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
        exceldenTabloyaOku(seciliExcelYolu, tableModel);
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

        // 🔹 Özet Paneli (YATAY)
        JPanel summaryPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 20, 10));
        summaryPanel.setBorder(javax.swing.BorderFactory.createTitledBorder("⚡ Filtre Sonuç Özeti"));

        JLabel lblToplamUmsatz = new JLabel("Umsatz: €0.00");
        lblToplamUmsatz.setForeground(java.awt.Color.BLUE);

        JLabel lblToplamTanken = new JLabel("Tanken: €0.00");
        lblToplamTanken.setForeground(java.awt.Color.RED);

        JLabel lblToplamSonstig = new JLabel("Sonstig: €0.00");
        lblToplamSonstig.setForeground(java.awt.Color.RED);

        JLabel lblToplamTeslim = new JLabel("Übergabe: €0.00");
        lblToplamTeslim.setForeground(java.awt.Color.GREEN.darker());

        JLabel lblToplamVorschuss = new JLabel("Vorschuss: €0.00");
        lblToplamVorschuss.setForeground(java.awt.Color.MAGENTA.darker());

        Font ozetFont = new Font("Segoe UI", Font.BOLD, 13);
        for (JLabel lbl : new JLabel[]{lblToplamUmsatz, lblToplamTanken, lblToplamSonstig, lblToplamTeslim, lblToplamVorschuss}) {
            lbl.setFont(ozetFont);
            summaryPanel.add(lbl);
        }

        btnFiltrele.addActionListener(e -> {
            yenidenTabloyuYukle(); // 📌 Tüm veriyi yüklemeden önce tablo temizlenir ve tekrar doldurulur

            java.util.Date basDate = baslangicChooser.getDate();
            java.util.Date bitDate = bitisChooser.getDate();
            String secilenPlaka = cbPlaka.getSelectedItem().toString();
            String secilenSurucu = cbSurucu.getSelectedItem().toString();

            if (basDate == null || bitDate == null) {
                JOptionPane.showMessageDialog(frame, "Lütfen geçerli bir tarih aralığı seçiniz.");
                return;
            }

            java.text.SimpleDateFormat df = new java.text.SimpleDateFormat("dd.MM.yyyy");

            // 🔽 FİLTRELEME
            for (int i = tableModel.getRowCount() - 1; i >= 0; i--) {
                try {
                    String tarihStr = tableModel.getValueAt(i, 2).toString();
                    java.util.Date satirTarihi = df.parse(tarihStr);

                    String plaka = tableModel.getValueAt(i, 1).toString();
                    String surucu = tableModel.getValueAt(i, 0).toString();

                    boolean tarihUygun = !satirTarihi.before(basDate) && !satirTarihi.after(bitDate);
                    boolean plakaUygun = secilenPlaka.equals("Tümü") || plaka.equals(secilenPlaka);
                    boolean surucuUygun = secilenSurucu.equals("Tümü") || surucu.equals(secilenSurucu);

                    if (!(tarihUygun && plakaUygun && surucuUygun)) {
                        tableModel.removeRow(i);
                    }
                } catch (Exception ex) {
                    tableModel.removeRow(i);
                }
            }

            // 🔽 ÖZET DEĞERLERİ HESAPLA
            double toplamUmsatz = 0, toplamTanken = 0, toplamSonstig = 0, toplamTeslim = 0, toplamVorschuss = 0;

            for (int i = 0; i < tableModel.getRowCount(); i++) {
                try {
                    String umsatzStr = tableModel.getValueAt(i, 10).toString().replace(",", ".").trim();
                    String tankenStr = tableModel.getValueAt(i, 7).toString().replace(",", ".").trim();
                    String sonstigStr = tableModel.getValueAt(i, 8).toString().replace(",", ".").trim();
                    String teslimStr = tableModel.getValueAt(i, 14).toString().replace(",", ".").trim();
                    String vorschussStr = tableModel.getValueAt(i, 15).toString().replace(",", ".").trim();

                    toplamUmsatz += Double.parseDouble(umsatzStr);
                    toplamTanken += Double.parseDouble(tankenStr);
                    toplamSonstig += Double.parseDouble(sonstigStr);
                    toplamTeslim += Double.parseDouble(teslimStr);
                    toplamVorschuss += Double.parseDouble(vorschussStr);
                } catch (Exception ignored) {}
            }

            lblToplamUmsatz.setText("Umsatz: €" + String.format("%.2f", toplamUmsatz));
            lblToplamTanken.setText("Tanken: €" + String.format("%.2f", toplamTanken));
            lblToplamSonstig.setText("Sonstig: €" + String.format("%.2f", toplamSonstig));
            lblToplamTeslim.setText("Übergabe: €" + String.format("%.2f", toplamTeslim));
            lblToplamVorschuss.setText("Vorschuss: €" + String.format("%.2f", toplamVorschuss));
        });


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
