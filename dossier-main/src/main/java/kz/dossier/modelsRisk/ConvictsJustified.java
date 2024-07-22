package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "convicts_justified", schema = "imp_risk")
public class ConvictsJustified {
    @Nullable
    private String iin;
    @Nullable

    private String last_name;
    @Nullable

    private String first_name;
    @Nullable
    private String patronomyc;
    @Nullable
    private String birth_date;
    @Nullable
    private String reg_date;
    @Nullable
    private String qualification;

    @Nullable
    private String severity_code_crime;
    @Nullable
    private String qualification_code;
    @Nullable
    private String measure_punishment;
    @Nullable
    private String code_desicion_by_person;
    @Nullable
    private String decision_on_person;
    @Nullable
    private String court_of_first_instance;
    @Nullable
    @Id
    private String erdr_number;
    @Nullable
    private String consider_date_first_instance;
    @Nullable
    private String code_started_investiogation;
    @Nullable
    private String investigative_authority;
    @Nullable
    private boolean is_iin_upd;
    @Nullable
    private boolean is_rab;
    @Nullable
    private Long id;

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getLast_name() {
        return last_name;
    }

    public void setLast_name(@Nullable String last_name) {
        this.last_name = last_name;
    }

    @Nullable
    public String getFirst_name() {
        return first_name;
    }

    public void setFirst_name(@Nullable String first_name) {
        this.first_name = first_name;
    }

    @Nullable
    public String getPatronomyc() {
        return patronomyc;
    }

    public void setPatronomyc(@Nullable String patronomyc) {
        this.patronomyc = patronomyc;
    }

    @Nullable
    public String getBirth_date() {
        return birth_date;
    }

    public void setBirth_date(@Nullable String birth_date) {
        this.birth_date = birth_date;
    }

    @Nullable
    public String getReg_date() {
        return reg_date;
    }

    public void setReg_date(@Nullable String reg_date) {
        this.reg_date = reg_date;
    }

    @Nullable
    public String getQualification() {
        return qualification;
    }

    public void setQualification(@Nullable String qualification) {
        this.qualification = qualification;
    }

    @Nullable
    public String getSeverity_code_crime() {
        return severity_code_crime;
    }

    public void setSeverity_code_crime(@Nullable String severity_code_crime) {
        this.severity_code_crime = severity_code_crime;
    }

    @Nullable
    public String getQualification_code() {
        return qualification_code;
    }

    public void setQualification_code(@Nullable String qualification_code) {
        this.qualification_code = qualification_code;
    }

    @Nullable
    public String getMeasure_punishment() {
        return measure_punishment;
    }

    public void setMeasure_punishment(@Nullable String measure_punishment) {
        this.measure_punishment = measure_punishment;
    }

    @Nullable
    public String getCode_desicion_by_person() {
        return code_desicion_by_person;
    }

    public void setCode_desicion_by_person(@Nullable String code_desicion_by_person) {
        this.code_desicion_by_person = code_desicion_by_person;
    }

    @Nullable
    public String getDecision_on_person() {
        return decision_on_person;
    }

    public void setDecision_on_person(@Nullable String decision_on_person) {
        this.decision_on_person = decision_on_person;
    }

    @Nullable
    public String getCourt_of_first_instance() {
        return court_of_first_instance;
    }

    public void setCourt_of_first_instance(@Nullable String court_of_first_instance) {
        this.court_of_first_instance = court_of_first_instance;
    }

    @Nullable
    public String getErdr_number() {
        return erdr_number;
    }

    public void setErdr_number(@Nullable String erdr_number) {
        this.erdr_number = erdr_number;
    }

    @Nullable
    public String getConsider_date_first_instance() {
        return consider_date_first_instance;
    }

    public void setConsider_date_first_instance(@Nullable String consider_date_first_instance) {
        this.consider_date_first_instance = consider_date_first_instance;
    }

    @Nullable
    public String getCode_started_investiogation() {
        return code_started_investiogation;
    }

    public void setCode_started_investiogation(@Nullable String code_started_investiogation) {
        this.code_started_investiogation = code_started_investiogation;
    }

    @Nullable
    public String getInvestigative_authority() {
        return investigative_authority;
    }

    public void setInvestigative_authority(@Nullable String investigative_authority) {
        this.investigative_authority = investigative_authority;
    }

    public boolean isIs_iin_upd() {
        return is_iin_upd;
    }

    public void setIs_iin_upd(boolean is_iin_upd) {
        this.is_iin_upd = is_iin_upd;
    }

    public boolean isIs_rab() {
        return is_rab;
    }

    public void setIs_rab(boolean is_rab) {
        this.is_rab = is_rab;
    }

    @Nullable
    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }
}
