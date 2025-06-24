package GUİ;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.SwingConstants;

public class logın extends JFrame {

    /**
     * 
     */
    private static final long serialVersionUID = 1L;
    private JTextField usernameField;
    private JPasswordField passwordField;

    public logın() {
        setTitle("Admin Anmeldung");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(500, 350);
        setLocationRelativeTo(null); 
        getContentPane().setLayout(null);

        JPanel panel = new JPanel();
        panel.setBackground(new Color(245, 245, 245));
        panel.setBounds(50, 50, 380, 220);
        panel.setLayout(null);
        getContentPane().add(panel);

        JLabel titleLabel = new JLabel("Fahrkontrollsystem", SwingConstants.CENTER);
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 18));
        titleLabel.setBounds(60, 10, 260, 30);
        panel.add(titleLabel);

        JLabel userLabel = new JLabel("Benutzername:");
        userLabel.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        userLabel.setBounds(40, 60, 100, 25);
        panel.add(userLabel);

        usernameField = new JTextField();
        usernameField.setBounds(150, 60, 180, 25);
        panel.add(usernameField);

        JLabel passLabel = new JLabel("Passwort:");
        passLabel.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        passLabel.setBounds(40, 100, 100, 25);
        panel.add(passLabel);

        passwordField = new JPasswordField();
        passwordField.setBounds(150, 100, 180, 25);
        panel.add(passwordField);

        JButton loginButton = new JButton("Anmelden");
        loginButton.setBackground(new Color(100, 149, 237));
        loginButton.setForeground(Color.WHITE);
        loginButton.setFocusPainted(false);
        loginButton.setFont(new Font("Segoe UI", Font.BOLD, 14));
        loginButton.setBounds(130, 150, 120, 30);
        panel.add(loginButton);

        loginButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String username = usernameField.getText().trim();
                String password = new String(passwordField.getPassword());

                if (username.equals("admin") && password.equals("1234")) {
                    JOptionPane.showMessageDialog(logın.this, "Anmeldung erfolgreich!", "Erfolg", JOptionPane.INFORMATION_MESSAGE);
                    dispose(); 
                    new anamenü(); 
                } else {
                    JOptionPane.showMessageDialog(logın.this, "Benutzername oder Passwort falsch!", "Fehler", JOptionPane.ERROR_MESSAGE);
                }
            }
        });
    }
}
