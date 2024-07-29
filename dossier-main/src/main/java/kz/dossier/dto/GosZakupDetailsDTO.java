package kz.dossier.dto;

public class GosZakupDetailsDTO {
    private String contractNumber;
    private String dealType;
    private String dealStatus;
    private String amount;
    private String date;
    private String methodOfBuying;
    private String nameOfOpposite;
    private String oppositeBin;

    public String getContractNumber() {
        return contractNumber;
    }

    public void setContractNumber(String contractNumber) {
        this.contractNumber = contractNumber;
    }

    public String getDealType() {
        return dealType;
    }

    public void setDealType(String dealType) {
        this.dealType = dealType;
    }

    public String getDealStatus() {
        return dealStatus;
    }

    public void setDealStatus(String dealStatus) {
        this.dealStatus = dealStatus;
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
