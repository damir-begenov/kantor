package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "imp_risk.mz")
public class MzEntity {
    @Id
    @Nullable
    private Long id;

    @Nullable
    @Column(name = "iin")
    private String iin;

    @Nullable
    @Column(name = "disease_code")
    private String diseaseCode;

    @Nullable
    @Column(name = "reg")
    private String reg;

    @Nullable
    @Column(name = "status_mz")
    private String statusMz;

    @Nullable
    @Column(name = "medical_org")
    private String medicalOrg;

    @Nullable
    public Long getId() {

        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getDiseaseCode() {
        return diseaseCode;
    }

    public void setDiseaseCode(@Nullable String diseaseCode) {
        this.diseaseCode = diseaseCode;
    }

    @Nullable
    public String getReg() {
        return reg;
    }

    public void setReg(@Nullable String reg) {
        this.reg = reg;
    }

    @Nullable
    public String getStatusMz() {
        return statusMz;
    }

    public void setStatusMz(@Nullable String statusMz) {
        this.statusMz = statusMz;
    }

    @Nullable
    public String getMedicalOrg() {
        return medicalOrg;
    }

    public void setMedicalOrg(@Nullable String medicalOrg) {
        this.medicalOrg = medicalOrg;
    }
}