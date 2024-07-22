package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "imp_kfm_ul_old.mv_reg_address_ul")
public class RegAddressUlEntity {
    @Nullable
    @Column(name = "subject_id")
    private Integer subjectId;
    @Nullable

    @Column(name = "layer_id")
    private Integer layerId;
    @Nullable

    @Column(name = "bin")
    private String bin;
    @Nullable

    @Column(name = "reg_action_code")
    private Integer regActionCode;
    @Nullable

    @Column(name = "reg_date")
    @Id

    private Date regDate;
    @Nullable

    @Column(name = "d_org_form")
    private Integer dOrgForm;
    @Nullable

    @Column(name = "d_org_legal_form")
    private Integer dOrgLegalForm;
    @Nullable

    @Column(name = "org_name_ru")
    private String orgNameRu;
    @Nullable

    @Column(name = "org_name_kz")
    private String orgNameKz;
    @Nullable

    @Column(name = "is_active")
    private Boolean isActive;
    @Nullable

    @Column(name = "reg_addr_region_ru")
    private String regAddrRegionRu;
    @Nullable

    @Column(name = "reg_addr_region_kz")
    private String regAddrRegionKz;
    @Nullable

    @Column(name = "reg_addr_district_ru")
    private String regAddrDistrictRu;
    @Nullable

    @Column(name = "reg_addr_district_kz")
    private String regAddrDistrictKz;
    @Nullable

    @Column(name = "reg_addr_rural_district_ru")
    private String regAddrRuralDistrictRu;
    @Nullable

    @Column(name = "reg_addr_rural_district_kz")
    private String regAddrRuralDistrictKz;
    @Nullable

    @Column(name = "reg_addr_locality_ru")
    private String regAddrLocalityRu;
    @Nullable

    @Column(name = "reg_addr_locality_kz")
    private String regAddrLocalityKz;
    @Nullable

    @Column(name = "reg_addr_street_ru")
    private String regAddrStreetRu;

    @Nullable
    @Column(name = "reg_addr_street_kz")
    private String regAddrStreetKz;

    @Nullable
    @Column(name = "reg_addr_bulding_num")
    private String regAddrBuildingNum;

    @Nullable
    @Column(name = "reg_addr_block_num")
    private String regAddrBlockNum;

    @Nullable
    @Column(name = "reg_addr_builing_body_num")
    private String regAddrBuildingBodyNum;

    @Nullable
    @Column(name = "reg_addr_office")
    private String regAddrOffice;

    @Nullable
    @Column(name = "oked_name_ru")
    private String okedNameRu;

    @Nullable
    @Column(name = "oked_name_kz")
    private String okedNameKz;

    @Nullable
    @Column(name = "authorized_capital_total")
    private String authorizedCapitalTotal;

    @Nullable
    @Column(name = "authorized_capital_gov_share")
    private String authorizedCapitalGovShare;

    @Nullable
    @Column(name = "authorized_capital_private_share")
    private String authorizedCapitalPrivateShare;

    @Nullable
    @Column(name = "ul_org_reg_number")
    private String ulOrgRegNumber;

    @Nullable
    @Column(name = "tax_inactive")
    private Boolean taxInactive;

    @Nullable
    @Column(name = "d_ul_formation_type")
    private Integer dUlFormationType;

    @Nullable
    @Column(name = "d_ownership_type")
    private String dOwnershipType;

    @Nullable
    @Column(name = "is_commercial")
    private Boolean isCommercial;

    @Nullable
    @Column(name = "is_entrepreneurial")
    private Boolean isEntrepreneurial;

    @Nullable
    @Column(name = "d_private_enterprise_type")
    private Integer dPrivateEnterpriseType;

    @Nullable
    @Column(name = "is_subsidiary")
    private Boolean isSubsidiary;
    @Nullable
    private Boolean is_international;
    @Nullable
    private Boolean is_foreign_investors;
    @Nullable
    private Boolean is_stateless_person;
    @Nullable
    private Integer d_state_enterprise_type;
    @Nullable
    private Integer d_ul_fund_type;
    @Nullable
    private Boolean is_government_agency;
    @Nullable
    private String ul_status;

    @Nullable
    public Integer getSubjectId() {
        return subjectId;
    }

    public void setSubjectId(@Nullable Integer subjectId) {
        this.subjectId = subjectId;
    }

    @Nullable
    public Integer getLayerId() {
        return layerId;
    }

    public void setLayerId(@Nullable Integer layerId) {
        this.layerId = layerId;
    }

    @Nullable
    public String getBin() {
        return bin;
    }

    public void setBin(@Nullable String bin) {
        this.bin = bin;
    }

    @Nullable
    public Integer getRegActionCode() {
        return regActionCode;
    }

    public void setRegActionCode(@Nullable Integer regActionCode) {
        this.regActionCode = regActionCode;
    }

    @Nullable
    public Date getRegDate() {
        return regDate;
    }

    public void setRegDate(@Nullable Date regDate) {
        this.regDate = regDate;
    }

    @Nullable
    public Integer getdOrgForm() {
        return dOrgForm;
    }

    public void setdOrgForm(@Nullable Integer dOrgForm) {
        this.dOrgForm = dOrgForm;
    }

    @Nullable
    public Integer getdOrgLegalForm() {
        return dOrgLegalForm;
    }

    public void setdOrgLegalForm(@Nullable Integer dOrgLegalForm) {
        this.dOrgLegalForm = dOrgLegalForm;
    }

    @Nullable
    public String getOrgNameRu() {
        return orgNameRu;
    }

    public void setOrgNameRu(@Nullable String orgNameRu) {
        this.orgNameRu = orgNameRu;
    }

    @Nullable
    public String getOrgNameKz() {
        return orgNameKz;
    }

    public void setOrgNameKz(@Nullable String orgNameKz) {
        this.orgNameKz = orgNameKz;
    }

    @Nullable
    public Boolean getActive() {
        return isActive;
    }

    public void setActive(@Nullable Boolean active) {
        isActive = active;
    }

    @Nullable
    public String getRegAddrRegionRu() {
        return regAddrRegionRu;
    }

    public void setRegAddrRegionRu(@Nullable String regAddrRegionRu) {
        this.regAddrRegionRu = regAddrRegionRu;
    }

    @Nullable
    public String getRegAddrRegionKz() {
        return regAddrRegionKz;
    }

    public void setRegAddrRegionKz(@Nullable String regAddrRegionKz) {
        this.regAddrRegionKz = regAddrRegionKz;
    }

    @Nullable
    public String getRegAddrDistrictRu() {
        return regAddrDistrictRu;
    }

    public void setRegAddrDistrictRu(@Nullable String regAddrDistrictRu) {
        this.regAddrDistrictRu = regAddrDistrictRu;
    }

    @Nullable
    public String getRegAddrDistrictKz() {
        return regAddrDistrictKz;
    }

    public void setRegAddrDistrictKz(@Nullable String regAddrDistrictKz) {
        this.regAddrDistrictKz = regAddrDistrictKz;
    }

    @Nullable
    public String getRegAddrRuralDistrictRu() {
        return regAddrRuralDistrictRu;
    }

    public void setRegAddrRuralDistrictRu(@Nullable String regAddrRuralDistrictRu) {
        this.regAddrRuralDistrictRu = regAddrRuralDistrictRu;
    }

    @Nullable
    public String getRegAddrRuralDistrictKz() {
        return regAddrRuralDistrictKz;
    }

    public void setRegAddrRuralDistrictKz(@Nullable String regAddrRuralDistrictKz) {
        this.regAddrRuralDistrictKz = regAddrRuralDistrictKz;
    }

    @Nullable
    public String getRegAddrLocalityRu() {
        return regAddrLocalityRu;
    }

    public void setRegAddrLocalityRu(@Nullable String regAddrLocalityRu) {
        this.regAddrLocalityRu = regAddrLocalityRu;
    }

    @Nullable
    public String getRegAddrLocalityKz() {
        return regAddrLocalityKz;
    }

    public void setRegAddrLocalityKz(@Nullable String regAddrLocalityKz) {
        this.regAddrLocalityKz = regAddrLocalityKz;
    }

    @Nullable
    public String getRegAddrStreetRu() {
        return regAddrStreetRu;
    }

    public void setRegAddrStreetRu(@Nullable String regAddrStreetRu) {
        this.regAddrStreetRu = regAddrStreetRu;
    }

    @Nullable
    public String getRegAddrStreetKz() {
        return regAddrStreetKz;
    }

    public void setRegAddrStreetKz(@Nullable String regAddrStreetKz) {
        this.regAddrStreetKz = regAddrStreetKz;
    }

    @Nullable
    public String getRegAddrBuildingNum() {
        return regAddrBuildingNum;
    }

    public void setRegAddrBuildingNum(@Nullable String regAddrBuildingNum) {
        this.regAddrBuildingNum = regAddrBuildingNum;
    }

    @Nullable
    public String getRegAddrBlockNum() {
        return regAddrBlockNum;
    }

    public void setRegAddrBlockNum(@Nullable String regAddrBlockNum) {
        this.regAddrBlockNum = regAddrBlockNum;
    }

    @Nullable
    public String getRegAddrBuildingBodyNum() {
        return regAddrBuildingBodyNum;
    }

    public void setRegAddrBuildingBodyNum(@Nullable String regAddrBuildingBodyNum) {
        this.regAddrBuildingBodyNum = regAddrBuildingBodyNum;
    }

    @Nullable
    public String getRegAddrOffice() {
        return regAddrOffice;
    }

    public void setRegAddrOffice(@Nullable String regAddrOffice) {
        this.regAddrOffice = regAddrOffice;
    }

    @Nullable
    public String getOkedNameRu() {
        return okedNameRu;
    }

    public void setOkedNameRu(@Nullable String okedNameRu) {
        this.okedNameRu = okedNameRu;
    }

    @Nullable
    public String getOkedNameKz() {
        return okedNameKz;
    }

    public void setOkedNameKz(@Nullable String okedNameKz) {
        this.okedNameKz = okedNameKz;
    }

    @Nullable
    public String getAuthorizedCapitalTotal() {
        return authorizedCapitalTotal;
    }

    public void setAuthorizedCapitalTotal(@Nullable String authorizedCapitalTotal) {
        this.authorizedCapitalTotal = authorizedCapitalTotal;
    }

    @Nullable
    public String getAuthorizedCapitalGovShare() {
        return authorizedCapitalGovShare;
    }

    public void setAuthorizedCapitalGovShare(@Nullable String authorizedCapitalGovShare) {
        this.authorizedCapitalGovShare = authorizedCapitalGovShare;
    }

    @Nullable
    public String getAuthorizedCapitalPrivateShare() {
        return authorizedCapitalPrivateShare;
    }

    public void setAuthorizedCapitalPrivateShare(@Nullable String authorizedCapitalPrivateShare) {
        this.authorizedCapitalPrivateShare = authorizedCapitalPrivateShare;
    }

    @Nullable
    public String getUlOrgRegNumber() {
        return ulOrgRegNumber;
    }

    public void setUlOrgRegNumber(@Nullable String ulOrgRegNumber) {
        this.ulOrgRegNumber = ulOrgRegNumber;
    }

    @Nullable
    public Boolean getTaxInactive() {
        return taxInactive;
    }

    public void setTaxInactive(@Nullable Boolean taxInactive) {
        this.taxInactive = taxInactive;
    }

    @Nullable
    public Integer getdUlFormationType() {
        return dUlFormationType;
    }

    public void setdUlFormationType(@Nullable Integer dUlFormationType) {
        this.dUlFormationType = dUlFormationType;
    }

    @Nullable
    public String getdOwnershipType() {
        return dOwnershipType;
    }

    public void setdOwnershipType(@Nullable String dOwnershipType) {
        this.dOwnershipType = dOwnershipType;
    }

    @Nullable
    public Boolean getCommercial() {
        return isCommercial;
    }

    public void setCommercial(@Nullable Boolean commercial) {
        isCommercial = commercial;
    }

    @Nullable
    public Boolean getEntrepreneurial() {
        return isEntrepreneurial;
    }

    public void setEntrepreneurial(@Nullable Boolean entrepreneurial) {
        isEntrepreneurial = entrepreneurial;
    }

    @Nullable
    public Integer getdPrivateEnterpriseType() {
        return dPrivateEnterpriseType;
    }

    public void setdPrivateEnterpriseType(@Nullable Integer dPrivateEnterpriseType) {
        this.dPrivateEnterpriseType = dPrivateEnterpriseType;
    }

    @Nullable
    public Boolean getSubsidiary() {
        return isSubsidiary;
    }

    public void setSubsidiary(@Nullable Boolean subsidiary) {
        isSubsidiary = subsidiary;
    }

    @Nullable
    public Boolean getIs_international() {
        return is_international;
    }

    public void setIs_international(@Nullable Boolean is_international) {
        this.is_international = is_international;
    }

    @Nullable
    public Boolean getIs_foreign_investors() {
        return is_foreign_investors;
    }

    public void setIs_foreign_investors(@Nullable Boolean is_foreign_investors) {
        this.is_foreign_investors = is_foreign_investors;
    }

    @Nullable
    public Boolean getIs_stateless_person() {
        return is_stateless_person;
    }

    public void setIs_stateless_person(@Nullable Boolean is_stateless_person) {
        this.is_stateless_person = is_stateless_person;
    }

    @Nullable
    public Integer getD_state_enterprise_type() {
        return d_state_enterprise_type;
    }

    public void setD_state_enterprise_type(@Nullable Integer d_state_enterprise_type) {
        this.d_state_enterprise_type = d_state_enterprise_type;
    }

    @Nullable
    public Integer getD_ul_fund_type() {
        return d_ul_fund_type;
    }

    public void setD_ul_fund_type(@Nullable Integer d_ul_fund_type) {
        this.d_ul_fund_type = d_ul_fund_type;
    }

    @Nullable
    public Boolean getIs_government_agency() {
        return is_government_agency;
    }

    public void setIs_government_agency(@Nullable Boolean is_government_agency) {
        this.is_government_agency = is_government_agency;
    }

    @Nullable
    public String getUl_status() {
        return ul_status;
    }

    public void setUl_status(@Nullable String ul_status) {
        this.ul_status = ul_status;
    }
}