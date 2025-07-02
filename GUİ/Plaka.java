package GUİ;

import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;

public class Plaka {
    private String plaka;
    private int sonKm;

    public Plaka(String plaka, int sonKm) {
        this.plaka = plaka;
        this.sonKm = sonKm;
    }

    public String getPlaka() {
        return plaka;
    }

    public void setPlaka(String plaka) {
        this.plaka = plaka;
    }

    public int getSonKm() {
        return sonKm;
    }

    public void setSonKm(int sonKm) {
        this.sonKm = sonKm;
    }

    public int getKalanKm() {
        int mod = sonKm % 15000;
        return (mod == 0) ? 15000 : 15000 - mod;
    }

    public boolean isBakimsiz() {
        return (sonKm % 15000 == 0);
    }

    public boolean muayeneZamaniMi(String secilenAy) {
        String bugununAyi = LocalDate.now().getMonth().getDisplayName(TextStyle.FULL, new Locale("tr", "TR"));
        return bugununAyi.equalsIgnoreCase(secilenAy) && isBakimsiz();
    }

    public Object[] toObjectArray() {
        return new Object[] {
            plaka,
            sonKm,
            getKalanKm() + " km",
            isBakimsiz() ? "BAKIMSIZ" : "BAKIMLI"
        };
    }
}
