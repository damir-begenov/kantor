package kz.dossier.modelsRisk;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "convicts_terminated_by_rehab", schema = "imp_risk")
public class ConvictsTerminatedByRehab {
    @Id
    private String erdr_number;
    private String iin;
    private String qualification_desc;
    private String qualification_code;
    private String qualificatiob_code_by_11;
    private String qualification_by_11;
    private String code_last_solution;
    private String last_solution;
    private String last_solution_date;
    private String last_name;
    private String first_name;
    private String patronomyc;
    private String birth_date;
    private String investigative_authority_code;
    private String investigative_authority;
    private String is_find_iin;
    private boolean is_rab;
    private Long id;

    public String getErdr_number() {
        return erdr_number;
    }

    public void setErdr_number(String erdr_number) {
        this.erdr_number = erdr_number;
    }

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public String getQualification_desc() {
        return qualification_desc;
    }

    public void setQualification_desc(String qualification_desc) {
        this.qualification_desc = qualification_desc;
    }

    public String getQualification_code() {
        return qualification_code;
    }

    public void setQualification_code(String qualification_code) {
        this.qualification_code = qualification_code;
    }

    public String getQualificatiob_code_by_11() {
        return qualificatiob_code_by_11;
    }

    public void setQualificatiob_code_by_11(String qualificatiob_code_by_11) {
        this.qualificatiob_code_by_11 = qualificatiob_code_by_11;
    }

    public String getQualification_by_11() {
        return qualification_by_11;
    }

    public void setQualification_by_11(String qualification_by_11) {
        this.qualification_by_11 = qualification_by_11;
    }

    public String getCode_last_solution() {
        return code_last_solution;
    }

    public void setCode_last_solution(String code_last_solution) {
        this.code_last_solution = code_last_solution;
    }

    public String getLast_solution() {
        return last_solution;
    }

    public void setLast_solution(String last_solution) {
        this.last_solution = last_solution;
    }

    public String getLast_solution_date() {
        return last_solution_date;
    }

    public void setLast_solution_date(String last_solution_date) {
        this.last_solution_date = last_solution_date;
    }

    public String getLast_name() {
        return last_name;
    }

    public void setLast_name(String last_name) {
        this.last_name = last_name;
    }

    public String getFirst_name() {
        return first_name;
    }

    public void setFirst_name(String first_name) {
        this.first_name = first_name;
    }

    public String getPatronomyc() {
        return patronomyc;
    }

    public void setPatronomyc(String patronomyc) {
        this.patronomyc = patronomyc;
    }

    public String getBirth_date() {
        return birth_date;
    }

    public void setBirth_date(String birth_date) {
        this.birth_date = birth_date;
    }

    public String getInvestigative_authority_code() {
        return investigative_authority_code;
    }

    public void setInvestigative_authority_code(String investigative_authority_code) {
        this.investigative_authority_code = investigative_authority_code;
    }

    public String getInvestigative_authority() {
        return investigative_authority;
    }

    public void setInvestigative_authority(String investigative_authority) {
        this.investigative_authority = investigative_authority;
    }

    public String getIs_find_iin() {
        return is_find_iin;
    }

    public void setIs_find_iin(String is_find_iin) {
        this.is_find_iin = is_find_iin;
    }

    public boolean isIs_rab() {
        return is_rab;
    }

    public void setIs_rab(boolean is_rab) {
        this.is_rab = is_rab;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }
}
