package kz.dossier.dto;

public class GosZakupDTO {
    private String period;
    private String sum;
    private Long number;
    private Long opposite;

    public String getPeriod() {
        return period;
    }

    public void setPeriod(String period) {
        this.period = period;
    }

    public String getSum() {
        return sum;
    }

    public void setSum(String sum) {
        this.sum = sum;
    }

    public Long getNumber() {
        return number;
    }

    public void setNumber(Long number) {
        this.number = number;
    }

    public Long getOpposite() {
        return opposite;
    }

    public void setOpposite(Long opposite) {
        this.opposite = opposite;
    }
}
