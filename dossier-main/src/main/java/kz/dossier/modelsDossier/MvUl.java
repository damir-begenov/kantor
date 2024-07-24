package kz.dossier.modelsDossier;

import jakarta.persistence.Transient;
import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "mv_ul", schema = "imp_kfm_ul")
public class MvUl {
    @Nullable
    private String bin;
    @Nullable
    private String full_name_rus;
    @Nullable
    private String full_name_kaz;
    @Nullable
    private String org_status;
    @Nullable
    private String org_registration_type;
    @Nullable
    private String org_form;
    @Nullable
    private String org_reg_date;
    @Nullable
    private String legal_form;
    @Nullable
    private String is_resident;
    @Nullable
    private String registration_agensy;
    @Nullable
    private String registration_number;
    @Nullable
    private String oked;
    @Nullable
    private String head_organization;
    @Nullable
    @Id
    private String subject_id;
    @Nullable
    private String layer_id;
    @Nullable
    private String ul_status;
    @Nullable
    @Transient
    private Boolean is_upd;
    @Nullable
    private String short_name;

    public String getBin() {
        return bin;
    }

    public void setBin(String bin) {
        this.bin = bin;
    }

    public String getFull_name_rus() {
        return full_name_rus;
    }

    public void setFull_name_rus(String full_name_rus) {
        this.full_name_rus = full_name_rus;
    }

    public String getFull_name_kaz() {
        return full_name_kaz;
    }

    public void setFull_name_kaz(String full_name_kaz) {
        this.full_name_kaz = full_name_kaz;
    }

    public String getOrg_status() {
        return org_status;
    }

    public void setOrg_status(String org_status) {
        this.org_status = org_status;
    }

    public String getOrg_registration_type() {
        return org_registration_type;
    }

    public void setOrg_registration_type(String org_registration_type) {
        this.org_registration_type = org_registration_type;
    }

    public String getOrg_form() {
        return org_form;
    }

    public void setOrg_form(String org_form) {
        this.org_form = org_form;
    }

    public String getOrg_reg_date() {
        return org_reg_date;
    }

    public void setOrg_reg_date(String org_reg_date) {
        this.org_reg_date = org_reg_date;
    }

    public String getLegal_form() {
        return legal_form;
    }

    public void setLegal_form(String legal_form) {
        this.legal_form = legal_form;
    }

    public String getIs_resident() {
        return is_resident;
    }

    public void setIs_resident(String is_resident) {
        this.is_resident = is_resident;
    }

    public String getRegistration_agensy() {
        return registration_agensy;
    }

    public void setRegistration_agensy(String registration_agensy) {
        this.registration_agensy = registration_agensy;
    }

    public String getRegistration_number() {
        return registration_number;
    }

    public void setRegistration_number(String registration_number) {
        this.registration_number = registration_number;
    }

    public String getOked() {
        return oked;
    }

    public void setOked(String oked) {
        this.oked = oked;
    }

    public String getHead_organization() {
        return head_organization;
    }

    public void setHead_organization(String head_organization) {
        this.head_organization = head_organization;
    }

    public String getSubject_id() {
        return subject_id;
    }

    public void setSubject_id(String subject_id) {
        this.subject_id = subject_id;
    }

    public String getLayer_id() {
        return layer_id;
    }

    public void setLayer_id(String layer_id) {
        this.layer_id = layer_id;
    }

    public String getUl_status() {
        return ul_status;
    }

    public void setUl_status(String ul_status) {
        this.ul_status = ul_status;
    }

    public Boolean isIs_upd() {
        return is_upd;
    }

    public void setIs_upd(Boolean is_upd) {
        this.is_upd = is_upd;
    }

    public String getShort_name() {
        return short_name;
    }

    public void setShort_name(String short_name) {
        this.short_name = short_name;
    }
}
