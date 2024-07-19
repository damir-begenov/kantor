package kz.dossier.modelsRisk;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.sql.Date;

@Entity
@Table(name = "block_esf", schema = "imp_risk")
public class BlockEsf {
    @Id
    private Long id;
    private String iin_bin;
    private Date start_dt;
    private Date end_dt;
    private Date update_dt;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getIin_bin() {
        return iin_bin;
    }

    public void setIin_bin(String iin_bin) {
        this.iin_bin = iin_bin;
    }

    public Date getStart_dt() {
        return start_dt;
    }

    public void setStart_dt(Date start_dt) {
        this.start_dt = start_dt;
    }

    public Date getEnd_dt() {
        return end_dt;
    }

    public void setEnd_dt(Date end_dt) {
        this.end_dt = end_dt;
    }

    public Date getUpdate_dt() {
        return update_dt;
    }

    public void setUpdate_dt(Date update_dt) {
        this.update_dt = update_dt;
    }
}
