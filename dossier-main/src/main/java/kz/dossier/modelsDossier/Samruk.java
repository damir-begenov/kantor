package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

import java.sql.Date;

@Entity
@Table(name = "samruk_kazyna", schema = "imp_kfm_db")
public class Samruk {
    @Id
    private Long contract_id;
    private String supplier;
    private String customer;
    private Double amount_with_nds;
    private Date contract_date;
    private String contract_status;
    private String contract_number;
    private Date date_of_conclusion;
    private String contract_type;
    private String purchase_method;

    public Long getContract_id() {
        return contract_id;
    }

    public void setContract_id(Long contract_id) {
        this.contract_id = contract_id;
    }

    public String getSupplier() {
        return supplier;
    }

    public void setSupplier(String supplier) {
        this.supplier = supplier;
    }

    public String getCustomer() {
        return customer;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public Double getAmount_with_nds() {
        return amount_with_nds;
    }

    public void setAmount_with_nds(Double amount_with_nds) {
        this.amount_with_nds = amount_with_nds;
    }

    public Date getContract_date() {
        return contract_date;
    }

    public void setContract_date(Date contract_date) {
        this.contract_date = contract_date;
    }

    public String getContract_status() {
        return contract_status;
    }

    public void setContract_status(String contract_status) {
        this.contract_status = contract_status;
    }

    public String getContract_number() {
        return contract_number;
    }

    public void setContract_number(String contract_number) {
        this.contract_number = contract_number;
    }

    public Date getDate_of_conclusion() {
        return date_of_conclusion;
    }

    public void setDate_of_conclusion(Date date_of_conclusion) {
        this.date_of_conclusion = date_of_conclusion;
    }

    public String getContract_type() {
        return contract_type;
    }

    public void setContract_type(String contract_type) {
        this.contract_type = contract_type;
    }

    public String getPurchase_method() {
        return purchase_method;
    }

    public void setPurchase_method(String purchase_method) {
        this.purchase_method = purchase_method;
    }
}
