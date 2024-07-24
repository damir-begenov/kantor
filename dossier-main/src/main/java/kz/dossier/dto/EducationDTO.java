package kz.dossier.dto;

public class EducationDTO {
    private String bin;
    private String nameOfPlace;
    private String typeOfPlace;
    private String startDate;
    private String endDate;
    private String specializationOrClass;
    private String difference;

    public String getBin() {
        return bin;
    }
    public String getDifference() {
        return difference;
    }
    public String getEndDate() {
        return endDate;
    }
    public String getNameOfPlace() {
        return nameOfPlace;
    }
    public String getSpecializationOrClass() {
        return specializationOrClass;
    }
    public String getStartDate() {
        return startDate;
    }
    public String getTypeOfPlace() {
        return typeOfPlace;
    }
    public void setBin(String bin) {
        this.bin = bin;
    }
    public void setDifference(String difference) {
        this.difference = difference;
    }
    public void setEndDate(String endDate) {
        this.endDate = endDate;
    }
    public void setNameOfPlace(String nameOfPlace) {
        this.nameOfPlace = nameOfPlace;
    }
    public void setSpecializationOrClass(String specializationOrClass) {
        this.specializationOrClass = specializationOrClass;
    }
    public void setStartDate(String startDate) {
        this.startDate = startDate;
    }
    public void setTypeOfPlace(String typeOfPlace) {
        this.typeOfPlace = typeOfPlace;
    }
    
}
