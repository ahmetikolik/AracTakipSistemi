package GUİ;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;

public class ButtonEditor extends DefaultCellEditor {
    protected JButton button;
    private String label;
    private boolean clicked;
    private JTable table;
    private DefaultTableModel model;

    public ButtonEditor(JCheckBox checkBox, JTable table, DefaultTableModel model) {
        super(checkBox);
        this.table = table;
        this.model = model;

        button = new JButton();
        button.setOpaque(true);

        button.addActionListener(e -> fireEditingStopped());
    }

    @Override
    public Component getTableCellEditorComponent(
            JTable table, Object value, boolean isSelected, int row, int column) {
        label = (value == null) ? "✖" : value.toString();
        button.setText(label);
        clicked = true;
        return button;
    }

    @Override
    public Object getCellEditorValue() {
        if (clicked) {
            int response = JOptionPane.showConfirmDialog(
                button, "Bu satırı silmek istediğinize emin misiniz?", "Onay",
                JOptionPane.YES_NO_OPTION);
            if (response == JOptionPane.YES_OPTION) {
                model.removeRow(table.getSelectedRow());

                // Excel'e de yansıt
                try {
                    ExcelExporter.exportToExcel(model, "surusler.xlsx");
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(button, "Excel'e silme yansıyamadı: " + ex.getMessage());
                }
            }
        }
        clicked = false;
        return label;
    }

    @Override
    public boolean stopCellEditing() {
        clicked = false;
        return super.stopCellEditing();
    }
}
