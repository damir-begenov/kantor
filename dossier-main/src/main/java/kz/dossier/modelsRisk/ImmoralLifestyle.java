package kz.dossier.modelsRisk;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.sql.Date;
import java.util.UUID;

@Entity
@Table(name = "immoral_lifestyle", schema = "imp_kfm_fl")
public class ImmoralLifestyle {
    @Id
    private Long id;
    private String authority_detected;
    private String iin;
    private String citizenship_id;
    private Date decision_date;
    private Integer fine_amount;
    private UUID mv_fl_iin_id;
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getAuthority_detected() {
        return authority_detected;
    }

    public void setAuthority_detected(String authority_detected) {
        this.authority_detected = authority_detected;
    }

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getCitizenship_id() {
        return citizenship_id;
    }

    public void setCitizenship_id(String citizenship_id) {
        this.citizenship_id = citizenship_id;
    }

    public Date getDecision_date() {
        return decision_date;
    }

    public void setDecision_date(Date decision_date) {
        this.decision_date = decision_date;
    }

    public Integer getFine_amount() {
        return fine_amount;
    }

    public void setFine_amount(Integer fine_amount) {
        this.fine_amount = fine_amount;
    }

    public UUID getMv_fl_iin_id() {
        return mv_fl_iin_id;
    }

    public void setMv_fl_iin_id(UUID mv_fl_iin_id) {
        this.mv_fl_iin_id = mv_fl_iin_id;
    }
}
