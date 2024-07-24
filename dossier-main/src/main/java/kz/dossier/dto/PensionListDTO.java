package kz.dossier.dto;

public class PensionListDTO {
    private String bin;
    private String name;
    private String period;
    private double sum010;
    private double sum012;

    public void setBin(String bin) {
        this.bin = bin;
    }
    public void setName(String name) {
        this.name = name;
    }
    public void setPeriod(String period) {
        this.period = period;
    }
    public void setSum010(double sum010) {
        this.sum010 = sum010;
    }
    public void setSum012(double sum012) {
        this.sum012 = sum012;
    }
    public String getBin() {
        return bin;
    }
    public String getName() {
        return name;
    }
    public String getPeriod() {
        return period;
    }
    public double getSum010() {
        return sum010;
    }
    public double getSum012() {
        return sum012;
    }
}
