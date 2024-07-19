package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "imp_risk.opg")
public class OpgEntity {
    @Nullable
    @Column(name = "region")
    private String region;

    @Nullable
    @Column(name = "groupe")
    private String groupe;
    @Nullable

    @Column(name = "name")
    private String name;
    @Id
    @Nullable

    @Column(name = "bin")
    private String bin;

    @Nullable
    @Column(name = "year")
    private Integer year;

    @Nullable
    public String getRegion() {
        return region;
    }

    public void setRegion(@Nullable String region) {
        this.region = region;
    }

    @Nullable
    public String getGroupe() {
        return groupe;
    }

    public void setGroupe(@Nullable String groupe) {
        this.groupe = groupe;
    }

    @Nullable
    public String getName() {
        return name;
    }

    public void setName(@Nullable String name) {
        this.name = name;
    }

    @Nullable
    public String getBin() {
        return bin;
    }

    public void setBin(@Nullable String bin) {
        this.bin = bin;
    }

    @Nullable
    public Integer getYear() {
        return year;
    }

    public void setYear(@Nullable Integer year) {
        this.year = year;
    }
}