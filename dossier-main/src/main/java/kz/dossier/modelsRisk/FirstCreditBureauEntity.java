package kz.dossier.modelsRisk;

import org.jetbrains.annotations.Nullable;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "imp_kfm_fl.first_credit_bureau")
public class FirstCreditBureauEntity {
    @Id
    @Nullable
    private Long id;

    @Nullable
    @Column(name = "type")
    private String type;

    @Nullable
    @Column(name = "credit_in_foid")
    private Double creditInFoid;

    @Nullable
    @Column(name = "iin")
    private String iin;

    @Column(name = "region")
    @Nullable
    private String region;

    @Column(name = "quantity_fpd_spd")
    @Nullable
    private Integer quantityFpdSpd;

    @Column(name = "amount_of_debt")
    @Nullable
    private Double amountOfDebt;

    @Column(name = "max_delay_day_num_1")
    @Nullable
    private Integer maxDelayDayNum1;

    @Column(name = "fin_institutions_name")
    @Nullable
    private String finInstitutionsName;

    @Column(name = "total_count_of_credits")
    @Nullable
    private Integer totalCountOfCredits;

    @Column(name = "total_sum_of_credits")
    @Nullable
    private Double totalSumOfCredits;

    @Column(name = "max_delay_day_num_2")
    @Nullable
    private Integer maxDelayDayNum2;

    @Column(name = "relevance_date")
    @Nullable
    private Date relevanceDate;

    @Nullable
    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }

    @Nullable
    public String getType() {
        return type;
    }

    public void setType(@Nullable String type) {
        this.type = type;
    }

    @Nullable
    public Double getCreditInFoid() {
        return creditInFoid;
    }

    public void setCreditInFoid(@Nullable Double creditInFoid) {
        this.creditInFoid = creditInFoid;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getRegion() {
        return region;
    }

    public void setRegion(@Nullable String region) {
        this.region = region;
    }

    @Nullable
    public Integer getQuantityFpdSpd() {
        return quantityFpdSpd;
    }

    public void setQuantityFpdSpd(@Nullable Integer quantityFpdSpd) {
        this.quantityFpdSpd = quantityFpdSpd;
    }

    @Nullable
    public Double getAmountOfDebt() {
        return amountOfDebt;
    }

    public void setAmountOfDebt(@Nullable Double amountOfDebt) {
        this.amountOfDebt = amountOfDebt;
    }

    @Nullable
    public Integer getMaxDelayDayNum1() {
        return maxDelayDayNum1;
    }

    public void setMaxDelayDayNum1(@Nullable Integer maxDelayDayNum1) {
        this.maxDelayDayNum1 = maxDelayDayNum1;
    }

    @Nullable
    public String getFinInstitutionsName() {
        return finInstitutionsName;
    }

    public void setFinInstitutionsName(@Nullable String finInstitutionsName) {
        this.finInstitutionsName = finInstitutionsName;
    }

    @Nullable
    public Integer getTotalCountOfCredits() {
        return totalCountOfCredits;
    }

    public void setTotalCountOfCredits(@Nullable Integer totalCountOfCredits) {
        this.totalCountOfCredits = totalCountOfCredits;
    }

    @Nullable
    public Double getTotalSumOfCredits() {
        return totalSumOfCredits;
    }

    public void setTotalSumOfCredits(@Nullable Double totalSumOfCredits) {
        this.totalSumOfCredits = totalSumOfCredits;
    }

    @Nullable
    public Integer getMaxDelayDayNum2() {
        return maxDelayDayNum2;
    }

    public void setMaxDelayDayNum2(@Nullable Integer maxDelayDayNum2) {
        this.maxDelayDayNum2 = maxDelayDayNum2;
    }

    @Nullable
    public Date getRelevanceDate() {
        return relevanceDate;
    }

    public void setRelevanceDate(@Nullable Date relevanceDate) {
        this.relevanceDate = relevanceDate;
    }
}
