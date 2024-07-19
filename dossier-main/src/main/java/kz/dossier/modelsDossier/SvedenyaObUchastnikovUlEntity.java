package kz.dossier.modelsDossier;

public class SvedenyaObUchastnikovUlEntity {
    private String iin_bin;
    private String FIOorUlName;
    private String identificator;
    private String reg_date;
    private String risk;

    public String getIin_bin() {
        return iin_bin;
    }

    public void setIin_bin(String iin_bin) {
        this.iin_bin = iin_bin;
    }

    public String getFIOorUlName() {
        return FIOorUlName;
    }

    public void setFIOorUlName(String FIOorUlName) {
        this.FIOorUlName = FIOorUlName;
    }

    public String getIdentificator() {
        return identificator;
    }

    public void setIdentificator(String identificator) {
        this.identificator = identificator;
    }

    public String getReg_date() {
        return reg_date;
    }

    public void setReg_date(String reg_date) {
        this.reg_date = reg_date;
    }

    public String getRisk() {
        return risk;
    }

    public void setRisk(String risk) {
        this.risk = risk;
    }
}
