package kz.dossier.dto;

import java.util.List;

public class SamrukKazynaForAll {
    private List<SamrukDTO> whenCustomer;
    private List<SamrukDTO> whenSupplier;

    public List<SamrukDTO> getWhenCustomer() {
        return whenCustomer;
    }

    public void setWhenCustomer(List<SamrukDTO> whenCustomer) {
        this.whenCustomer = whenCustomer;
    }

    public List<SamrukDTO> getWhenSupplier() {
        return whenSupplier;
    }

    public void setWhenSupplier(List<SamrukDTO> whenSupplier) {
        this.whenSupplier = whenSupplier;
    }
}
