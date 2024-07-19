package kz.dossier.modelsRisk;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "criminals", schema = "imp_risk")
public class Criminals {
    private String court_name;
    private String court_dt;
    private String iin_bin;
    private String decision;
    private String crime_name;
    private String crime_code;
    private String sentence;
    private String add_info;
    private String treatment;
    private String erdr;
    private Date update_dt;
    @Id
    private long id;

    public String getCourt_name() {
        return court_name;
    }

    public void setCourt_name(String court_name) {
        this.court_name = court_name;
    }

    public String getCourt_dt() {
        return court_dt;
    }

    public void setCourt_dt(String court_dt) {
        this.court_dt = court_dt;
    }

    public String getIin_bin() {
        return iin_bin;
    }

    public void setIin_bin(String iin_bin) {
        this.iin_bin = iin_bin;
    }

    public String getDecision() {
        return decision;
    }

    public void setDecision(String decision) {
        this.decision = decision;
    }

    public String getCrime_name() {
        return crime_name;
    }

    public void setCrime_name(String crime_name) {
        this.crime_name = crime_name;
    }

    public String getCrime_code() {
        return crime_code;
    }

    public void setCrime_code(String crime_code) {
        this.crime_code = crime_code;
    }

    public String getSentence() {
        return sentence;
    }

    public void setSentence(String sentence) {
        this.sentence = sentence;
    }

    public String getAdd_info() {
        return add_info;
    }

    public void setAdd_info(String add_info) {
        this.add_info = add_info;
    }

    public String getTreatment() {
        return treatment;
    }

    public void setTreatment(String treatment) {
        this.treatment = treatment;
    }

    public String getErdr() {
        return erdr;
    }

    public void setErdr(String erdr) {
        this.erdr = erdr;
    }

    public Date getUpdate_dt() {
        return update_dt;
    }

    public void setUpdate_dt(Date update_dt) {
        this.update_dt = update_dt;
    }

    public long getId() {
        return id;
    }

    public void setId(long id) {
        this.id = id;
    }
}
