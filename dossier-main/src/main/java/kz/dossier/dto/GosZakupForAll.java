package kz.dossier.dto;

import java.util.List;

public class GosZakupForAll {
    private List<GosZakupDTO> whenSupplier;
    private List<GosZakupDTO> whenCustomer;

    public List<GosZakupDTO> getWhenSupplier() {
        return whenSupplier;
    }

    public void setWhenSupplier(List<GosZakupDTO> whenSupplier) {
        this.whenSupplier = whenSupplier;
    }

    public List<GosZakupDTO> getWhenCustomer() {
        return whenCustomer;
    }

    public void setWhenCustomer(List<GosZakupDTO> whenCustomer) {
        this.whenCustomer = whenCustomer;
    }
}
