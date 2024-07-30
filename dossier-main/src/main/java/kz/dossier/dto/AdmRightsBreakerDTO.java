package kz.dossier.dto;

public class AdmRightsBreakerDTO {
    private String iin;
    private String organ;
    private String dateOfStart;
    private String numberOfProtocol;
    private String placeOfWork;
    private String qualification;
    private String forcedImplementation;
    private String periodTo;
    private String sumOfFine;
    private String reasonOfStopping;

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getOrgan() {
        return organ;
    }

    public void setOrgan(String organ) {
        this.organ = organ;
    }

    public String getDateOfStart() {
        return dateOfStart;
    }

    public void setDateOfStart(String dateOfStart) {
        this.dateOfStart = dateOfStart;
    }

    public String getNumberOfProtocol() {
        return numberOfProtocol;
    }

    public void setNumberOfProtocol(String numberOfProtocol) {
        this.numberOfProtocol = numberOfProtocol;
    }

    public String getPlaceOfWork() {
        return placeOfWork;
    }

    public void setPlaceOfWork(String placeOfWork) {
        this.placeOfWork = placeOfWork;
    }

    public String getQualification() {
        return qualification;
    }

    public void setQualification(String qualification) {
        this.qualification = qualification;
    }

    public String getForcedImplementation() {
        return forcedImplementation;
    }

    public void setForcedImplementation(String forcedImplementation) {
        this.forcedImplementation = forcedImplementation;
    }

    public String getPeriodTo() {
        return periodTo;
    }

    public void setPeriodTo(String periodTo) {
        this.periodTo = periodTo;
    }

    public String getSumOfFine() {
        return sumOfFine;
    }

    public void setSumOfFine(String sumOfFine) {
        this.sumOfFine = sumOfFine;
    }

    public String getReasonOfStopping() {
        return reasonOfStopping;
    }

    public void setReasonOfStopping(String reasonOfStopping) {
        this.reasonOfStopping = reasonOfStopping;
    }
}
