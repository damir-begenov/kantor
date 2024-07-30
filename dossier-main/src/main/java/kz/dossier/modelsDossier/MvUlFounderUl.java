package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "imp_kfm_ul.ul_founder_ul")
public class MvUlFounderUl {

    @Id
    @Nullable
    private String id;
    @Nullable
    @Column(name = "reg_date")
    private String regDate;
    @Nullable

    @Column(name = "bin_org")
    private String binOrg;
    @Nullable

    @Column(name = "is_resident")
    private boolean isResident;
    @Nullable

    @Column(name = "founder_bin")
    private String founderBin;
    @Nullable

    @Column(name = "founder_reg_date")
    private String founderRegDate;
    @Nullable

    @Column(name = "founder_name_ru")
    private String founderNameRu;
    @Nullable

    @Column(name = "deposit")
    private String deposit;
    @Nullable

    @Column(name = "\"share\"")
    private String share;
    @Nullable

    @Column(name = "is_curr")
    private boolean isCurrent;

    @Nullable
    public String getId() {
        return id;
    }

    public void setId(@Nullable String id) {
        this.id = id;
    }

    @Nullable
    public String getRegDate() {
        return regDate;
    }

    public void setRegDate(@Nullable String regDate) {
        this.regDate = regDate;
    }

    @Nullable
    public String getBinOrg() {
        return binOrg;
    }

    public void setBinOrg(@Nullable String binOrg) {
        this.binOrg = binOrg;
    }

    public boolean isResident() {
        return isResident;
    }

    public void setResident(boolean resident) {
        isResident = resident;
    }

    @Nullable
    public String getFounderBin() {
        return founderBin;
    }

    public void setFounderBin(@Nullable String founderBin) {
        this.founderBin = founderBin;
    }

    @Nullable
    public String getFounderRegDate() {
        return founderRegDate;
    }

    public void setFounderRegDate(@Nullable String founderRegDate) {
        this.founderRegDate = founderRegDate;
    }

    @Nullable
    public String getFounderNameRu() {
        return founderNameRu;
    }

    public void setFounderNameRu(@Nullable String founderNameRu) {
        this.founderNameRu = founderNameRu;
    }

    @Nullable
    public String getDeposit() {
        return deposit;
    }

    public void setDeposit(@Nullable String deposit) {
        this.deposit = deposit;
    }

    @Nullable
    public String getShare() {
        return share;
    }

    public void setShare(@Nullable String share) {
        this.share = share;
    }

    public boolean isCurrent() {
        return isCurrent;
    }

    public void setCurrent(boolean current) {
        isCurrent = current;
    }
}