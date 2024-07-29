package kz.dossier.modelsDossier;

import jakarta.persistence.*;

@Entity
@Table(name = "imp_notary.advocate_list")
public class AdvocateListEntity {
    @Id
    @Column(name = "iin")
    private String iin;

    @Column(name = "fio")
    private String fio;

    @Column(name = "license_date")
    private String licenseDate;

    @Column(name = "license_number")
    private String licenseNumber;

    @Column(name = "status")
    private String status;

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getFio() {
        return fio;
    }

    public void setFio(String fio) {
        this.fio = fio;
    }

    public String getLicenseDate() {
        return licenseDate;
    }

    public void setLicenseDate(String licenseDate) {
        this.licenseDate = licenseDate;
    }

    public String getLicenseNumber() {
        return licenseNumber;
    }

    public void setLicenseNumber(String licenseNumber) {
        this.licenseNumber = licenseNumber;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }
}