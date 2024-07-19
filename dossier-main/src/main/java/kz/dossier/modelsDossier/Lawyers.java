package kz.dossier.modelsDossier;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "lawyers", schema = "imp_kfm_fl")
public class Lawyers {
    @Id
    private Long id;
    private String layers_fio;
    private String entry_date;
    private String date_of_issue;
    private String contacts;
    private String region;
    @Column(name = "№")
    private String number;
    private String license_num;
    private String address;
    private String iin;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getLayers_fio() {
        return layers_fio;
    }

    public void setLayers_fio(String layers_fio) {
        this.layers_fio = layers_fio;
    }

    public String getEntry_date() {
        return entry_date;
    }

    public void setEntry_date(String entry_date) {
        this.entry_date = entry_date;
    }

    public String getDate_of_issue() {
        return date_of_issue;
    }

    public void setDate_of_issue(String date_of_issue) {
        this.date_of_issue = date_of_issue;
    }

    public String getContacts() {
        return contacts;
    }

    public void setContacts(String contacts) {
        this.contacts = contacts;
    }

    public String getRegion() {
        return region;
    }

    public void setRegion(String region) {
        this.region = region;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getLicense_num() {
        return license_num;
    }

    public void setLicense_num(String license_num) {
        this.license_num = license_num;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }
}
