package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;


@Entity
@Table(name = "pdl", schema = "imp_kfm_fl")
public class Pdl {
    @Id
    @Nullable
    private String id;
    @Nullable
    private String bin;
    @Nullable
    private String iin;
    @Nullable
    private String organization_fullname;
    @Nullable
    private String fio;
    private String organ;
    @Nullable
    private String oblast;
    @Nullable
    private String position;
    @Nullable
    private String spouse_fio;
    @Nullable
    private String spouse_organ;
    @Nullable
    private String spouse_position;
    @Nullable
    private String spouse_iin;
    @Nullable

    public String getId() {
        return id;
    }
    @Nullable

    public void setId(String id) {
        this.id = id;
    }
    @Nullable

    public String getBin() {
        return bin;
    }
    @Nullable

    public void setBin(String bin) {
        this.bin = bin;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    @Nullable
    public void setIin(String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getOrganization_fullname() {
        return organization_fullname;
    }
    @Nullable

    public void setOrganization_fullname(String organization_fullname) {
        this.organization_fullname = organization_fullname;
    }
    @Nullable

    public String getFio() {
        return fio;
    }

    @Nullable
    public void setFio(String fio) {
        this.fio = fio;
    }

    @Nullable
    public String getOrgan() {
        return organ;
    }

    @Nullable
    public void setOrgan(String organ) {
        this.organ = organ;
    }

    @Nullable
    public String getOblast() {
        return oblast;
    }
    @Nullable

    public void setOblast(String oblast) {
        this.oblast = oblast;
    }

    @Nullable
    public String getPosition() {
        return position;
    }

    @Nullable
    public void setPosition(String position) {
        this.position = position;
    }
    @Nullable

    public String getSpouse_fio() {
        return spouse_fio;
    }

    @Nullable
    public void setSpouse_fio(String spouse_fio) {
        this.spouse_fio = spouse_fio;
    }

    @Nullable
    public String getSpouse_organ() {
        return spouse_organ;
    }

    @Nullable
    public void setSpouse_organ(String spouse_organ) {
        this.spouse_organ = spouse_organ;
    }

    @Nullable
    public String getSpouse_position() {
        return spouse_position;
    }
    @Nullable

    public void setSpouse_position(String spouse_position) {
        this.spouse_position = spouse_position;
    }

    @Nullable
    public String getSpouse_iin() {
        return spouse_iin;
    }

    @Nullable
    public void setSpouse_iin(String spouse_iin) {
        this.spouse_iin = spouse_iin;
    }
}
