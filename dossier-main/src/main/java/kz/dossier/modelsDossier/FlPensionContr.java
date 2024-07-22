package kz.dossier.modelsDossier;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "fl_pension", schema = "fl_pension_contr")
@NoArgsConstructor
@AllArgsConstructor
public class FlPensionContr {




    @Nullable
    @Column(name = "SURNAME")
    private String surname;
    @Nullable

    @Column(name = "FIRSTNAME")
    private String firstname;

    @Nullable
    @Column(name = "SECONDNAME")
    private String secondname;

    @Nullable
    @Column(name = "IIN")
    private String iin;

    @Column(name = "PAY_DATE")
    @Nullable

    private String payDate;

    @Column(name = "P_NAME")
    @Nullable
    private String pName;

    @Column(name = "P_RNN")
    @Nullable
    private String pRnn;

    @Column(name = "KNP")
    @Nullable

    private String knp;
    @Column(name = "PERIOD")
    @Nullable
    private String period;

    @Nullable
    @Column(name = "AMOUNT")
    private Double amount;
    @Nullable

    @Column(name = "year")
    private String year;

    @Nullable @Column(name = "id")
    @Id

    private Long id;

    @Nullable
    public String getSurname() {
        return surname;
    }

    public void setSurname(@Nullable String surname) {
        this.surname = surname;
    }

    @Nullable
    public String getFirstname() {
        return firstname;
    }

    public void setFirstname(@Nullable String firstname) {
        this.firstname = firstname;
    }

    @Nullable
    public String getSecondname() {
        return secondname;
    }

    public void setSecondname(@Nullable String secondname) {
        this.secondname = secondname;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getPayDate() {
        return payDate;
    }

    public void setPayDate(@Nullable String payDate) {
        this.payDate = payDate;
    }

    @Nullable
    public String getpName() {
        return pName;
    }

    public void setpName(@Nullable String pName) {
        this.pName = pName;
    }

    @Nullable
    public String getpRnn() {
        return pRnn;
    }

    public void setpRnn(@Nullable String pRnn) {
        this.pRnn = pRnn;
    }

    @Nullable
    public String getKnp() {
        return knp;
    }

    public void setKnp(@Nullable String knp) {
        this.knp = knp;
    }

    @Nullable
    public String getPeriod() {
        return period;
    }

    public void setPeriod(@Nullable String period) {
        this.period = period;
    }

    @Nullable
    public Double getAmount() {
        return amount;
    }

    public void setAmount(@Nullable Double amount) {
        this.amount = amount;
    }

    @Nullable
    public String getYear() {
        return year;
    }

    public void setYear(@Nullable String year) {
        this.year = year;
    }

    @Nullable
    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }
}
