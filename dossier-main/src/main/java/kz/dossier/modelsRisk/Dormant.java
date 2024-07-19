package kz.dossier.modelsRisk;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "dormant", schema = "imp_risk")
public class Dormant {

    private String iin_bin;
    private String rnn;
    private String taxpayer_name;
    private String taxpayer_fio;
    private String leader_fio;
    private String leader_iin;
    private String leader_rnn;
    private String order_number;
    @Id

    private String order_date;
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

    public String getOrder_number() {
        return order_number;
    }

    public void setOrder_number(String order_number) {
        this.order_number = order_number;
    }

    public String getOrder_date() {
        return order_date;
    }

    public void setOrder_date(String order_date) {
        this.order_date = order_date;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }
}
