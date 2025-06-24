package GUİ;

import java.awt.Component;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.IOException;

class UpdateButtonEditor extends DefaultCellEditor {
    private JButton button;
    private JTable table;
    private DefaultTableModel model;
    private JFrame parentFrame;

    public UpdateButtonEditor(JCheckBox checkBox, JTable table, DefaultTableModel model, JFrame parentFrame) {
        super(checkBox);
        this.table = table;
        this.model = model;
        this.parentFrame = parentFrame;

        button = new JButton("✎");
        button.addActionListener(e -> {
            int row = table.getSelectedRow();
            if (row != -1) {
                // Güncellenecek satırdan verileri al
                String surucu = model.getValueAt(row, 0).toString();
                String plaka = model.getValueAt(row, 1).toString();
                String tarih = model.getValueAt(row, 2).toString();

                String basKm = model.getValueAt(row, 3).toString();
                String bitKm = model.getValueAt(row, 4).toString();
                String bar = model.getValueAt(row, 5).toString();
                String rf = model.getValueAt(row, 6).toString();
                String tanken = model.getValueAt(row, 7).toString();
                String sonstig = model.getValueAt(row, 8).toString();
                String cc = model.getValueAt(row, 9).toString();

                // Güncelleme için kullanıcıdan veri al
                JTextField tfBar = new JTextField(bar);
                JTextField tfRf = new JTextField(rf);
                JTextField tfTanken = new JTextField(tanken);
                JTextField tfSonstig = new JTextField(sonstig);
                JTextField tfCc = new JTextField(cc);

                JPanel inputPanel = new JPanel();
                inputPanel.setLayout(new BoxLayout(inputPanel, BoxLayout.Y_AXIS));
                inputPanel.add(new JLabel("BAR (€):")); inputPanel.add(tfBar);
                inputPanel.add(new JLabel("RF (€):")); inputPanel.add(tfRf);
                inputPanel.add(new JLabel("TANKEN (€):")); inputPanel.add(tfTanken);
                inputPanel.add(new JLabel("SONSTIG (€):")); inputPanel.add(tfSonstig);
                inputPanel.add(new JLabel("CC:")); inputPanel.add(tfCc);

                int result = JOptionPane.showConfirmDialog(parentFrame, inputPanel,
                        "Verileri Güncelle - " + surucu + " / " + plaka, JOptionPane.OK_CANCEL_OPTION);

                if (result == JOptionPane.OK_OPTION) {
                    try {
                        int barVal = Integer.parseInt(tfBar.getText().trim());
                        int rfVal = Integer.parseInt(tfRf.getText().trim());
                        int tankenVal = Integer.parseInt(tfTanken.getText().trim());
                        int sonstigVal = Integer.parseInt(tfSonstig.getText().trim());
                        int ccVal = Integer.parseInt(tfCc.getText().trim());

                        int basKmVal = Integer.parseInt(basKm);
                        int bitKmVal = Integer.parseInt(bitKm);
                        int gefKm = bitKmVal - basKmVal;
                        int umsatz = barVal + rfVal - sonstigVal + ccVal;
                        int diffKm = gefKm;
                        int dursKm = gefKm != 0 ? gefKm / 28 : 0;
                        int teslim = (int) (umsatz * 0.4);
                        int vorschuss = umsatz - teslim;

                        // Güncelle tablo satırını
                        model.setValueAt(barVal, row, 5);
                        model.setValueAt(rfVal, row, 6);
                        model.setValueAt(tankenVal, row, 7);
                        model.setValueAt(sonstigVal, row, 8);
                        model.setValueAt(ccVal, row, 9);

                        model.setValueAt(umsatz, row, 10);
                        model.setValueAt(diffKm, row, 11);
                        model.setValueAt(gefKm, row, 12);
                        model.setValueAt(dursKm, row, 13);
                        model.setValueAt(teslim, row, 14);
                        model.setValueAt(vorschuss, row, 15);

                        // Excel'e kaydet
                        ExcelExporter.exportToExcel(model, "surusler.xlsx");

                        JOptionPane.showMessageDialog(parentFrame, "Güncelleme başarılı.");
                    } catch (Exception ex) {
                        JOptionPane.showMessageDialog(parentFrame, "Hata: " + ex.getMessage());
                    }
                }
            }
        });
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value,
                                                 boolean isSelected, int row, int column) {
        return button;
    }

    @Override
    public Object getCellEditorValue() {
        return "";
    }
}
