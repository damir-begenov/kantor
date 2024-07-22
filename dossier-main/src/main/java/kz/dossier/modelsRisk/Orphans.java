package kz.dossier.modelsRisk;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "orphans", schema = "imp_risk")
public class Orphans {
    private String iin;
    private String region;

//    private String distict;
    private String fio;
    private String orphanage_name;
    private String graduation_year;
    private String birth_date;
    @Id

    private Integer id;

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getRegion() {
        return region;
    }

    public void setRegion(String region) {
        this.region = region;
    }

//    public String getDistict() {
//        return distict;
//    }
//
//    public void setDistict(String distict) {
//        this.distict = distict;
//    }

    public String getFio() {
        return fio;
    }

    public void setFio(String fio) {
        this.fio = fio;
    }

    public String getOrphanage_name() {
        return orphanage_name;
    }

    public void setOrphanage_name(String orphanage_name) {
        this.orphanage_name = orphanage_name;
    }

    public String getGraduation_year() {
        return graduation_year;
    }

    public void setGraduation_year(String graduation_year) {
        this.graduation_year = graduation_year;
    }

    public String getBirth_date() {
        return birth_date;
    }

    public void setBirth_date(String birth_date) {
        this.birth_date = birth_date;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }
}
