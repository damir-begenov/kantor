package kz.dossier.modelsDossier;

import kz.dossier.modelsRisk.*;
import kz.dossier.modelsRisk.OpgEntity;

import java.util.List;
import java.util.Map;

public class NodesUL {
    private List<MvUl> mvUls;
    private List<MvUlFounderFl> mvUlFounderFls;
    private List<AccountantListEntity> accountantListEntities;


    private List<Omn> omns; //risk
    private List<Equipment> equipment;
    private List<Dormant> dormants;
    private List<Bankrot> bankrots; //risk
    private List<OpgEntity> opgEntities; //risk

    private List<Adm> adms;  //risk

    public List<MvUlFounderFl> getMvUlFounderFls() {
        return mvUlFounderFls;
    }
    private List<Msh> mshes;
    private List<Criminals> criminals;  //risk
    private List<BlockEsf> blockEsfs;  //risk
    private List<NdsEntity> ndsEntities;
    private List<MvRnOld> mvRnOlds;
    private List<TaxOutEntity> taxOutEntities;
    private List<FpgTempEntity> fpgTempEntities; //Risk
    private List<Pdl> pdls;
    private List<kz.dossier.modelsDossier.QoldauSubsidy> QoldauSubsidy;
    private List<CommodityProducer> commodityProducers;
    private RegAddressUlEntity regAddressUlEntities;
    private List<SvedenyaObUchastnikovUlEntity> svedenyaObUchastnikovUlEntities;
    private List<FlContacts> fl_contacts;

    private Integer taxCount;

    private List<MvAutoFl> mvAutoFls;
    private List<FlPensionFinal> flPensionContrs;
    private List<Map<String, Object>> PensionYearAndEmpNum;
    private RegAddressUlEntity RegUlNaOdnomMeste;

    public RegAddressUlEntity getRegUlNaOdnomMeste() {
        return RegUlNaOdnomMeste;
    }

    public Integer getTaxCount() {
        return taxCount;
    }

    public void setTaxCount(Integer taxCount) {
        this.taxCount = taxCount;
    }

    public void setRegUlNaOdnomMeste(RegAddressUlEntity regUlNaOdnomMeste) {
        RegUlNaOdnomMeste = regUlNaOdnomMeste;
    }

    public List<Map<String, Object>> getPensionYearAndEmpNum() {
        return PensionYearAndEmpNum;
    }

    public void setPensionYearAndEmpNum(List<Map<String, Object>> pensionYearAndEmpNum) {
        PensionYearAndEmpNum = pensionYearAndEmpNum;
    }

    public List<FlPensionFinal> getFlPensionContrs() {
        return flPensionContrs;
    }

    public void setFlPensionContrs(List<FlPensionFinal> flPensionContrs) {
        this.flPensionContrs = flPensionContrs;
    }


    public List<FlContacts> getFl_contacts() {
        return fl_contacts;
    }

    public void setFl_contacts(List<FlContacts> fl_contacts) {
        this.fl_contacts = fl_contacts;
    }

    public List<MvAutoFl> getMvAutoFls() {
        return mvAutoFls;
    }

    public void setMvAutoFls(List<MvAutoFl> mvAutoFls) {
        this.mvAutoFls = mvAutoFls;
    }

    private boolean Person_with_risk;

    public boolean isPerson_with_risk() {
        return Person_with_risk;
    }

    public void setPerson_with_risk(boolean person_with_risk) {
        Person_with_risk = person_with_risk;
    }

    public List<SvedenyaObUchastnikovUlEntity> getSvedenyaObUchastnikovUlEntities() {
        return svedenyaObUchastnikovUlEntities;
    }

    public void setSvedenyaObUchastnikovUlEntities(List<SvedenyaObUchastnikovUlEntity> svedenyaObUchastnikovUlEntities) {
        this.svedenyaObUchastnikovUlEntities = svedenyaObUchastnikovUlEntities;
    }

    public RegAddressUlEntity getRegAddressUlEntities() {
        return regAddressUlEntities;
    }

    public void setRegAddressUlEntities(RegAddressUlEntity regAddressUlEntities) {
        this.regAddressUlEntities = regAddressUlEntities;
    }


    public List<CommodityProducer> getCommodityProducers() {
        return commodityProducers;
    }

    public void setCommodityProducers(List<CommodityProducer> commodityProducers) {
        this.commodityProducers = commodityProducers;
    }

    public List<QoldauSubsidy> getQoldauSubsidy() {
        return QoldauSubsidy;
    }

    public void setQoldauSubsidy(List<QoldauSubsidy> qoldauSubsidy) {
        QoldauSubsidy = qoldauSubsidy;
    }

    public List<Pdl> getPdls() {
        return pdls;
    }

    public void setPdls(List<Pdl> pdls) {
        this.pdls = pdls;
    }

    public List<FpgTempEntity> getFpgTempEntities() {
        return fpgTempEntities;
    }

    public void setFpgTempEntities(List<FpgTempEntity> fpgTempEntities) {
        this.fpgTempEntities = fpgTempEntities;
    }

    public List<TaxOutEntity> getTaxOutEntities() {
        return taxOutEntities;
    }

    public void setTaxOutEntities(List<TaxOutEntity> taxOutEntities) {
        this.taxOutEntities = taxOutEntities;
    }

    public List<MvRnOld> getMvRnOlds() {
        return mvRnOlds;
    }

    public void setMvRnOlds(List<MvRnOld> mvRnOlds) {
        this.mvRnOlds = mvRnOlds;
    }

    public List<NdsEntity> getNdsEntities() {
        return ndsEntities;
    }

    public void setNdsEntities(List<NdsEntity> ndsEntities) {
        this.ndsEntities = ndsEntities;
    }

    public List<BlockEsf> getBlockEsfs() {
        return blockEsfs;
    }

    public void setBlockEsfs(List<BlockEsf> blockEsfs) {
        this.blockEsfs = blockEsfs;
    }

    public List<OpgEntity> getOpgEntities() {
        return opgEntities;
    }

    public void setOpgEntities(List<OpgEntity> opgEntities) {
        this.opgEntities = opgEntities;
    }

    public List<Criminals> getCriminals() {
        return criminals;
    }

    public void setCriminals(List<Criminals> criminals) {
        this.criminals = criminals;
    }

    public List<Msh> getMshes() {
        return mshes;
    }

    public void setMshes(List<Msh> mshes) {
        this.mshes = mshes;
    }

    public List<AccountantListEntity> getAccountantListEntities() {
        return accountantListEntities;
    }

    public void setAccountantListEntities(List<AccountantListEntity> accountantListEntities) {
        this.accountantListEntities = accountantListEntities;
    }

    public void setMvUlFounderFls(List<MvUlFounderFl> mvUlFounderFls) {
        this.mvUlFounderFls = mvUlFounderFls;
    }

    public List<Bankrot> getBankrots() {
        return bankrots;
    }

    public void setBankrots(List<Bankrot> bankrots) {
        this.bankrots = bankrots;
    }

    public List<MvUl> getMvUls() {
        return mvUls;
    }

    public void setMvUls(List<MvUl> mvUls) {
        this.mvUls = mvUls;
    }

    public List<Omn> getOmns() {
        return omns;
    }

    public void setOmns(List<Omn> omns) {
        this.omns = omns;
    }

    public List<Equipment> getEquipment() {
        return equipment;
    }

    public void setEquipment(List<Equipment> equipment) {
        this.equipment = equipment;
    }

    public List<Dormant> getDormants() {
        return dormants;
    }

    public void setDormants(List<Dormant> dormants) {
        this.dormants = dormants;
    }

    public List<Adm> getAdms() {
        return adms;
    }

    public void setAdms(List<Adm> adms) {
        this.adms = adms;
    }
}
