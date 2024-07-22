package kz.dossier.modelsDossier;


import jakarta.persistence.Entity;
import jakarta.persistence.Id;

@Entity
public class FlRelativiesDTO {
    @Id
    private Long id;
    private String iin;
    private String fio;
    private String birth_date;
    private String parent_iin;
    private String parent_fio;
    private String parent_birth_date;
    private String relative_type;
    private String marriage_reg_date;
    private String marriage_divorce_date;
    private String level;

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

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

    public String getBirth_date() {
        return birth_date;
    }

    public void setBirth_date(String birth_date) {
        this.birth_date = birth_date;
    }

    public String getParent_iin() {
        return parent_iin;
    }

    public void setParent_iin(String parent_iin) {
        this.parent_iin = parent_iin;
    }

    public String getParent_fio() {
        return parent_fio;
    }

    public void setParent_fio(String parent_fio) {
        this.parent_fio = parent_fio;
    }

    public String getParent_birth_date() {
        return parent_birth_date;
    }

    public void setParent_birth_date(String parent_birth_date) {
        this.parent_birth_date = parent_birth_date;
    }

    public String getRelative_type() {
        return relative_type;
    }

    public void setRelative_type(String relative_type) {
        if(relative_type == "Супруг"){
            this.relative_type = "ЖЕНА";
        }else{
            this.relative_type = relative_type;
        }}

    public String getMarriage_reg_date() {
        return marriage_reg_date;
    }

    public void setMarriage_reg_date(String marriage_reg_date) {
        this.marriage_reg_date = marriage_reg_date;
    }

    public String getMarriage_divorce_date() {
        return marriage_divorce_date;
    }

    public void setMarriage_divorce_date(String marriage_divorce_date) {
        this.marriage_divorce_date = marriage_divorce_date;
    }
}
