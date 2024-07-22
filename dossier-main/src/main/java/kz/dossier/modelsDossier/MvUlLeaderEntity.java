package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "imp_kfm_ul_old.mv_ul_leader")
public class MvUlLeaderEntity {
    @Id
    @Nullable
    private String id;
    @Nullable

    @Column(name = "bin_org")
    private String binOrg;
    @Nullable

    @Column(name = "reg_date")
    private String regDate;
    @Nullable

    @Column(name = "iin")
    private String iin;
    @Nullable

    @Column(name = "lastname")
    private String lastName;
    @Nullable

    @Column(name = "firstname")
    private String firstName;
    @Nullable

    @Column(name = "patronymic")
    private String patronymic;
    @Nullable

    @Column(name = "position_id")
    private Integer positionId;
    @Nullable

    @Column(name = "appointment_date")
    private Date appointmentDate;
    @Nullable

    @Column(name = "removal_date")
    private Date removalDate;
    @Nullable

    @Column(name = "is_curr")
    private Boolean isCurrent;
    @Nullable

    @Column(name = "ul_status")
    private String ulStatus;

    @Nullable
    public String getId() {
        return id;
    }

    public void setId(@Nullable String id) {
        this.id = id;
    }

    @Nullable
    public String getBinOrg() {
        return binOrg;
    }

    public void setBinOrg(@Nullable String binOrg) {
        this.binOrg = binOrg;
    }

    @Nullable
    public String getRegDate() {
        return regDate;
    }

    public void setRegDate(@Nullable String regDate) {
        this.regDate = regDate;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getLastName() {
        return lastName;
    }

    public void setLastName(@Nullable String lastName) {
        this.lastName = lastName;
    }

    @Nullable
    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(@Nullable String firstName) {
        this.firstName = firstName;
    }

    @Nullable
    public String getPatronymic() {
        return patronymic;
    }

    public void setPatronymic(@Nullable String patronymic) {
        this.patronymic = patronymic;
    }

    @Nullable
    public Integer getPositionId() {
        return positionId;
    }

    public void setPositionId(@Nullable Integer positionId) {
        this.positionId = positionId;
    }

    @Nullable
    public Date getAppointmentDate() {
        return appointmentDate;
    }

    public void setAppointmentDate(@Nullable Date appointmentDate) {
        this.appointmentDate = appointmentDate;
    }

    @Nullable
    public Date getRemovalDate() {
        return removalDate;
    }

    public void setRemovalDate(@Nullable Date removalDate) {
        this.removalDate = removalDate;
    }

    @Nullable
    public Boolean getCurrent() {
        return isCurrent;
    }

    public void setCurrent(@Nullable Boolean current) {
        isCurrent = current;
    }

    @Nullable
    public String getUlStatus() {
        return ulStatus;
    }

    public void setUlStatus(@Nullable String ulStatus) {
        this.ulStatus = ulStatus;
    }
}