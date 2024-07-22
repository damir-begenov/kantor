package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.UUID;

@Entity
@Table(name = "ul_founder_fl", schema = "initial_data")

public class MvUlFounderFl {
    @Id
    private UUID id;
    private String reg_date;
    private String bin_org;
    private String iin;
    private String lastname;
    private String firstname;
    private String patronymic;
    private String deposit;
    private String share;
    private boolean is_curr;

    public UUID getId() {
        return id;
    }

    public void setId(UUID id) {
        this.id = id;
    }

    public String getReg_date() {
        return reg_date;
    }

    public void setReg_date(String reg_date) {
        this.reg_date = reg_date;
    }

    public boolean isIs_curr() {
        return is_curr;
    }

    public void setIs_curr(boolean is_curr) {
        this.is_curr = is_curr;
    }

    public String getBin_org() {
        return bin_org;
    }

    public void setBin_org(String bin_org) {
        this.bin_org = bin_org;
    }

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getLastname() {
        return lastname;
    }

    public void setLastname(String lastname) {
        this.lastname = lastname;
    }

    public String getFirstname() {
        return firstname;
    }

    public void setFirstname(String firstname) {
        this.firstname = firstname;
    }

    public String getPatronymic() {
        return patronymic;
    }

    public void setPatronymic(String patronymic) {
        this.patronymic = patronymic;
    }

    public String getDeposit() {
        return deposit;
    }

    public void setDeposit(String deposit) {
        this.deposit = deposit;
    }

    public String getShare() {
        return share;
    }

    public void setShare(String share) {
        this.share = share;
    }


}
