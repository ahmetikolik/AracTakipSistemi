package GUİ;

public class Plaka {
    private String plaka;
    private int baslangicKm;
    private int sonKm;
    private int bakimli; // 1 = bakımlı, 0 = bakımsız

    public Plaka(String plaka, int baslangicKm, int sonKm, int bakimli) {
        this.plaka = plaka;
        this.baslangicKm = baslangicKm;
        this.sonKm = sonKm;
        this.bakimli = bakimli;
    }

    public String getPlaka() {
        return plaka;
    }

    public int getBaslangicKm() {
        return baslangicKm;
    }

    public void setBaslangicKm(int baslangicKm) {
        this.baslangicKm = baslangicKm;
    }

    public int getSonKm() {
        return sonKm;
    }

    public void setSonKm(int sonKm) {
        this.sonKm = sonKm;
    }

    public int getBakimli() {
        return bakimli;
    }

    public void setBakimli(int bakimli) {
        this.bakimli = bakimli;
    }

    public void guncelleKm(int yeniSonKm) {
        this.sonKm = yeniSonKm;
        if ((sonKm - baslangicKm) >= 15000) {
            this.bakimli = 0;
            this.baslangicKm = sonKm; // bakım sonrası başlangıç güncellenir
        }
    }

    @Override
    public String toString() {
        return plaka + "," + baslangicKm + "," + sonKm + "," + bakimli;
    }
}
