package kz.dossier.modelsRisk;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "omn", schema = "imp_risk")
public class Omn {
    @Id
    private String iin_bin;
    private String rnn;
    private String taxpayer_name;
    private String taxpayer_fio;
    private String leader_fio;
    private String leader_iin;
    private String leader_rnn;
    private String decision_number;
    private String decision_date;
    private Integer id;

    public String getIin_bin() {
        return iin_bin;
    }

    public void setIin_bin(String iin_bin) {
        this.iin_bin = iin_bin;
    }

    public String getRnn() {
        return rnn;
    }

    public void setRnn(String rnn) {
        this.rnn = rnn;
    }

    public String getTaxpayer_name() {
        return taxpayer_name;
    }

    public void setTaxpayer_name(String taxpayer_name) {
        this.taxpayer_name = taxpayer_name;
    }

    public String getTaxpayer_fio() {
        return taxpayer_fio;
    }

    public void setTaxpayer_fio(String taxpayer_fio) {
        this.taxpayer_fio = taxpayer_fio;
    }

    public String getLeader_fio() {
        return leader_fio;
    }

    public void setLeader_fio(String leader_fio) {
        this.leader_fio = leader_fio;
    }

    public String getLeader_iin() {
        return leader_iin;
    }

    public void setLeader_iin(String leader_iin) {
        this.leader_iin = leader_iin;
    }

    public String getLeader_rnn() {
        return leader_rnn;
    }

    public void setLeader_rnn(String leader_rnn) {
        this.leader_rnn = leader_rnn;
    }

    public String getDecision_number() {
        return decision_number;
    }

    public void setDecision_number(String decision_number) {
        this.decision_number = decision_number;
    }

    public String getDecision_date() {
        return decision_date;
    }

    public void setDecision_date(String decision_date) {
        this.decision_date = decision_date;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }
}
