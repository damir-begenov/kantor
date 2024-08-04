package kz.dossier.dto;

public class ChangeFioDTO {
    private String historicalFIO;
    private String reasonOfChange;
    private String dateOfChange;

    public String getHistoricalFIO() {
        return historicalFIO;
    }

    public void setHistoricalFIO(String historicalFIO) {
        this.historicalFIO = historicalFIO;
    }

    public String getReasonOfChange() {
        return reasonOfChange;
    }

    public void setReasonOfChange(String reasonOfChange) {
        this.reasonOfChange = reasonOfChange;
    }

    public String getDateOfChange() {
        return dateOfChange;
    }

    public void setDateOfChange(String dateOfChange) {
        this.dateOfChange = dateOfChange;
    }
}
