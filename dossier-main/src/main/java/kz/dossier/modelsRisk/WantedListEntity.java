package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "imp_kfm_fl.wanted_list")
public class WantedListEntity {
    @Id
    @Nullable
    private Long id;

    @Nullable
    @Column(name = "iin")
    private String iin;
    @Nullable

    @Column(name = "days")
    private Integer days;
    @Nullable

    @Column(name = "organ")
    private String organ;
//    @Nullable

//    @Column(name = "initsiator")
//    private String initiator;
    @Nullable

    @Column(name = "status")
    private String status;
    @Nullable

    @Column(name = "relevance_date")
    private Date relevanceDate;

    @Nullable
    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public Integer getDays() {
        return days;
    }

    public void setDays(@Nullable Integer days) {
        this.days = days;
    }

    @Nullable
    public String getOrgan() {
        return organ;
    }

    public void setOrgan(@Nullable String organ) {
        this.organ = organ;
    }

//    @Nullable
//    public String getInitiator() {
//        return initiator;
//    }

//    public void setInitiator(@Nullable String initiator) {
//        this.initiator = initiator;
//    }

    @Nullable
    public String getStatus() {
        return status;
    }

    public void setStatus(@Nullable String status) {
        this.status = status;
    }

    @Nullable
    public Date getRelevanceDate() {
        return relevanceDate;
    }

    public void setRelevanceDate(@Nullable Date relevanceDate) {
        this.relevanceDate = relevanceDate;
    }
}