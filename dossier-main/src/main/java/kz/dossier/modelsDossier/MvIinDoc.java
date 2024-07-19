package kz.dossier.modelsDossier;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.sql.Date;
import java.util.UUID;

@Entity
@Table(name = "fl_iin_doc", schema = "initial_data")
public class MvIinDoc {
    @Id
    private UUID id;
    @Column
    private String gender;
    private Date birth_date;
    private String iin;
    private String citizenship_id;
    private String citizenship_ru_name;
    private String nationality_id;
    private String nationality_ru_name;
    private Boolean is_resident;
    private String life_status_id;
    private String life_status_ru_name;
    private Date death_date;
    private String doc_number;
    private String doc_type_id;
    private String doc_type_ru_name;
    private Date issue_date;
    private Date expiry_date;
    private String document_invalidity_id;
    private String document_invalidity_ru_name;
    private Date document_invalidity_date;
    private String issue_organization_id;
    private String issue_organization_ru_name;

    public UUID getId() {
        return id;
    }

    public void setId(UUID id) {
        this.id = id;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public Date getBirth_date() {
        return birth_date;
    }

    public void setBirth_date(Date birth_date) {
        this.birth_date = birth_date;
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

    public String getCitizenship_ru_name() {
        return citizenship_ru_name;
    }

    public void setCitizenship_ru_name(String citizenship_ru_name) {
        this.citizenship_ru_name = citizenship_ru_name;
    }

    public String getNationality_id() {
        return nationality_id;
    }

    public void setNationality_id(String nationality_id) {
        this.nationality_id = nationality_id;
    }

    public String getNationality_ru_name() {
        return nationality_ru_name;
    }

    public void setNationality_ru_name(String nationality_ru_name) {
        this.nationality_ru_name = nationality_ru_name;
    }

    public Boolean getIs_resident() {
        return is_resident;
    }

    public void setIs_resident(Boolean is_resident) {
        this.is_resident = is_resident;
    }

    public String getLife_status_id() {
        return life_status_id;
    }

    public void setLife_status_id(String life_status_id) {
        this.life_status_id = life_status_id;
    }

    public String getLife_status_ru_name() {
        return life_status_ru_name;
    }

    public void setLife_status_ru_name(String life_status_ru_name) {
        this.life_status_ru_name = life_status_ru_name;
    }

    public Date getDeath_date() {
        return death_date;
    }

    public void setDeath_date(Date death_date) {
        this.death_date = death_date;
    }

    public String getDoc_number() {
        return doc_number;
    }

    public void setDoc_number(String doc_number) {
        this.doc_number = doc_number;
    }

    public String getDoc_type_id() {
        return doc_type_id;
    }

    public void setDoc_type_id(String doc_type_id) {
        this.doc_type_id = doc_type_id;
    }

    public String getDoc_type_ru_name() {
        return doc_type_ru_name;
    }

    public void setDoc_type_ru_name(String doc_type_ru_name) {
        this.doc_type_ru_name = doc_type_ru_name;
    }

    public Date getIssue_date() {
        return issue_date;
    }

    public void setIssue_date(Date issue_date) {
        this.issue_date = issue_date;
    }

    public Date getExpiry_date() {
        return expiry_date;
    }

    public void setExpiry_date(Date expiry_date) {
        this.expiry_date = expiry_date;
    }

    public String getDocument_invalidity_id() {
        return document_invalidity_id;
    }

    public void setDocument_invalidity_id(String document_invalidity_id) {
        this.document_invalidity_id = document_invalidity_id;
    }

    public String getDocument_invalidity_ru_name() {
        return document_invalidity_ru_name;
    }

    public void setDocument_invalidity_ru_name(String document_invalidity_ru_name) {
        this.document_invalidity_ru_name = document_invalidity_ru_name;
    }

    public Date getDocument_invalidity_date() {
        return document_invalidity_date;
    }

    public void setDocument_invalidity_date(Date document_invalidity_date) {
        this.document_invalidity_date = document_invalidity_date;
    }

    public String getIssue_organization_id() {
        return issue_organization_id;
    }

    public void setIssue_organization_id(String issue_organization_id) {
        this.issue_organization_id = issue_organization_id;
    }

    public String getIssue_organization_ru_name() {
        return issue_organization_ru_name;
    }

    public void setIssue_organization_ru_name(String issue_organization_ru_name) {
        this.issue_organization_ru_name = issue_organization_ru_name;
    }
}
