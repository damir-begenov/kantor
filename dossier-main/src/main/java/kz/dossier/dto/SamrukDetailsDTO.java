package kz.dossier.dto;

public class SamrukDetailsDTO {
    private String numberOfDeal;
    private String typeOfDeal;
    private String statusOfDeal;
    private String amount;
    private String date;
    private String methodOfBuying;
    private String nameOfOpposite;
    private String oppositeBin;

    public String getNumberOfDeal() {
        return numberOfDeal;
    }

    public void setNumberOfDeal(String numberOfDeal) {
        this.numberOfDeal = numberOfDeal;
    }

    public String getTypeOfDeal() {
        return typeOfDeal;
    }

    public void setTypeOfDeal(String typeOfDeal) {
        this.typeOfDeal = typeOfDeal;
    }

    public String getStatusOfDeal() {
        return statusOfDeal;
    }

    public void setStatusOfDeal(String statusOfDeal) {
        this.statusOfDeal = statusOfDeal;
    }

    public String getAmount() {
        return amount;
    }

    public void setAmount(String amount) {
        this.amount = amount;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getMethodOfBuying() {
        return methodOfBuying;
    }

    public void setMethodOfBuying(String methodOfBuying) {
        this.methodOfBuying = methodOfBuying;
    }

    public String getNameOfOpposite() {
        return nameOfOpposite;
    }

    public void setNameOfOpposite(String nameOfOpposite) {
        this.nameOfOpposite = nameOfOpposite;
    }

    public String getOppositeBin() {
        return oppositeBin;
    }

    public void setOppositeBin(String oppositeBin) {
        this.oppositeBin = oppositeBin;
    }
}
