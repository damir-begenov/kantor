package kz.dossier.dto;

public class ULULMemberDTO {
    private String binIin;
    private String name;
    private String position;
    private String date;
    private Integer risksNumber;

    public String getBinIin() {
        return binIin;
    }

    public void setBinIin(String binIin) {
        this.binIin = binIin;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public Integer getRisksNumber() {
        return risksNumber;
    }

    public void setRisksNumber(Integer risksNumber) {
        this.risksNumber = risksNumber;
    }
}
