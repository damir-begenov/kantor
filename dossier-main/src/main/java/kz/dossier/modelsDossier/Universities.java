package kz.dossier.modelsDossier;


import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.sql.Date;

@Entity
@Table(name = "study", schema = "public")
public class Universities {
    @Id
    private Integer id;
    private String lastname;
    private String name;
    private String patronymic;
    private String iin;
    private Date start_date;
    @Column(name = "5")
    private String five;
    private String spec_name_2;
    private String course;
    private String duration;
    private String spec_name;
    private String study_name;
    private String study_type;
    private String study_code;
    private Date end_date;
    @Column(name = "14")
    private String fourteen;
    @Column(name = "15")
    private String fiveteen;
    @Column(name = "16")
    private String sixteen;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getLastname() {
        return lastname;
    }

    public void setLastname(String lastname) {
        this.lastname = lastname;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPatronymic() {
        return patronymic;
    }

    public void setPatronymic(String patronymic) {
        this.patronymic = patronymic;
    }

    public String getIin() {
        return iin;
    }

    public void setIin(String iin) {
        this.iin = iin;
    }

    public Date getStart_date() {
        return start_date;
    }

    public void setStart_date(Date start_date) {
        this.start_date = start_date;
    }

    public String getFive() {
        return five;
    }

    public void setFive(String five) {
        this.five = five;
    }

    public String getSpec_name_2() {
        return spec_name_2;
    }

    public void setSpec_name_2(String spec_name_2) {
        this.spec_name_2 = spec_name_2;
    }

    public String getCourse() {
        return course;
    }

    public void setCourse(String course) {
        this.course = course;
    }

    public String getDuration() {
        return duration;
    }

    public void setDuration(String duration) {
        this.duration = duration;
    }

    public String getSpec_name() {
        return spec_name;
    }

    public void setSpec_name(String spec_name) {
        this.spec_name = spec_name;
    }

    public String getStudy_name() {
        return study_name;
    }

    public void setStudy_name(String study_name) {
        this.study_name = study_name;
    }

    public String getStudy_type() {
        return study_type;
    }

    public void setStudy_type(String study_type) {
        this.study_type = study_type;
    }

    public String getStudy_code() {
        return study_code;
    }

    public void setStudy_code(String study_code) {
        this.study_code = study_code;
    }

    public Date getEnd_date() {
        return end_date;
    }

    public void setEnd_date(Date end_date) {
        this.end_date = end_date;
    }

    public String getFourteen() {
        return fourteen;
    }

    public void setFourteen(String fourteen) {
        this.fourteen = fourteen;
    }

    public String getFiveteen() {
        return fiveteen;
    }

    public void setFiveteen(String fiveteen) {
        this.fiveteen = fiveteen;
    }

    public String getSixteen() {
        return sixteen;
    }

    public void setSixteen(String sixteen) {
        this.sixteen = sixteen;
    }
}
