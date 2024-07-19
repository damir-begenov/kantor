package kz.dossier.modelsDossier;

import org.hibernate.annotations.Immutable;

import org.jetbrains.annotations.Nullable;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "fl_pension_contr", schema = "imp_kfm_fl")
@Immutable
public class FlPensionMini {


    @Column(name = "PAY_DATE")
    @Nullable

    private String payDate;

    @Column(name = "P_NAME")
    @Nullable
    @Id
    private String pName;

    @Column(name = "KNP")
    @Nullable
    private String knp;
    @Nullable

    @Column(name = "AMOUNT")
    private Double amount;

    @Override
    public String toString() {
        return "flPensionMini{" +
                "payDate='" + payDate + '\'' +
                ", pName='" + pName + '\'' +
                ", knp='" + knp + '\'' +
                ", amount=" + amount +
                '}';
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
    public String getKnp() {
        return knp;
    }

    public void setKnp(@Nullable String knp) {
        this.knp = knp;
    }

    @Nullable
    public Double getAmount() {
        return amount;
    }

    public void setAmount(@Nullable Double amount) {
        this.amount = amount;
    }
}
