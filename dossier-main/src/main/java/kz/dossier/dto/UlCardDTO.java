package kz.dossier.dto;

public class UlCardDTO {
    private String bin;
    private String name;
    private String status;
    private String regDate;

    public String getRegDate() {
        return regDate;
    }
    public void setRegDate(String regDate) {
        this.regDate = regDate;
    }

    public String getBin() {
        return bin;
    }
    public String getName() {
        return name;
    }
    public String getStatus() {
        return status;
    }
    public void setBin(String bin) {
        this.bin = bin;
    }
    public void setName(String name) {
        this.name = name;
    }
    public void setStatus(String status) {
        this.status = status;
    }

}
