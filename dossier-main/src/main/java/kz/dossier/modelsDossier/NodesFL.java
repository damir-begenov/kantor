package kz.dossier.modelsDossier;

import kz.dossier.modelsRisk.*;

import java.util.List;

public class NodesFL {
    private List<MvFl> mvFls;
    private List<PhotoDb> photoDbf;
    private List<MvAutoFl> mvAutoFls;
    private List<Omn> omns; //risk
    private List<DrugAddicts> drugAddicts; //risk
    private List<Incapacitated> incapacitateds; //risk
    private List<Kuis> kuis; //risk

    public List<Kuis> getKuis() {
        return kuis;
    }

    public void setKuis(List<Kuis> kuis) {
        this.kuis = kuis;
    }

    private List<Orphans> orphanss; //detdom
    private List<Equipment> equipment;
    private List<Dormant> dormants;
    private List<Adm> adms; //risk
    private List<FlRelativiesDTO> fl_relatives;
    private List<RegAddressFl> regAddressFls;

    private List<Criminals> criminals; //risk
    private List<ImmoralLifestyle> amoral; //risk
    private List<Dismissal> dismissals; //risk
    private List<BeneficiariesList> beneficiariesLists; //risk
    private List<ConvictsAbroad> convictsAbroads; //risk

    public List<ConvictsAbroad> getConvictsAbroads() {
        return convictsAbroads;
    }

    public void setConvictsAbroads(List<ConvictsAbroad> convictsAbroads) {
        this.convictsAbroads = convictsAbroads;
    }

    public List<Incapacitated> getIncapacitateds() {
        return incapacitateds;
    }

    public void setIncapacitateds(List<Incapacitated> incapacitateds) {
        this.incapacitateds = incapacitateds;
    }

    private List<Pdl> pdls;
    private List<WantedListEntity> wantedListEntities; //risk
    private List<CommodityProducer> commodityProducers;

    private List<MvIinDoc> mvIinDocs;
    private List<Universities> universities;

    private List<School> schools;
    private List<FlContacts> contacts;

    private List<MillitaryAccount> millitaryAccounts;
    private List<MilitaryAccounting2Entity> militaryAccounting2Entities;
    private List<ConvictsTerminatedByRehab> convictsTerminatedByRehabs; //risk
    private List<ConvictsJustified> convictsJustifieds;  //risk
    private List<Bankrot> bankrots; //risk
    private List<FlPensionFinal> flPensionContrs;
    private List<Msh> mshes;
    private List<BlockEsf> blockEsfs;  //risk
    private List<MvUlFounderFl> mvUlFounderFls;
    private List<NdsEntity> ndsEntities;
    private List<IpgoEmailEntity> ipgoEmailEntities;
    private List<kz.dossier.modelsDossier.TIpEntity>TIpEntity;
    private List<AccountantListEntity> accountantListEntities;
    private List<AdvocateListEntity> advocateListEntities;
    private List<AuditorsListEntity> auditorsListEntities;
    private List<BailiffListEntity> bailiffListEntities;
    private List<MzEntity> mzEntities ;
    private List<MvRnOld> mvRnOlds;


    private List<MvUlLeader> ul_leaderList; //Сведения об участии в ЮЛ

    public List<MvUlLeader> getUl_leaderList() {
        return ul_leaderList;
    }

    public void setUl_leaderList(List<MvUlLeader> ul_leaderList) {
        this.ul_leaderList = ul_leaderList;
    }

    public List<CommodityProducer> getCommodityProducers() {
        return commodityProducers;
    }
    public void setCommodityProducers(List<CommodityProducer> commodityProducers) {
        this.commodityProducers = commodityProducers;
    }

    public List<WantedListEntity> getWantedListEntities() {
        return wantedListEntities;
    }

    public void setWantedListEntities(List<WantedListEntity> wantedListEntities) {
        this.wantedListEntities = wantedListEntities;
    }

    public List<NdsEntity> getNdsEntities() {
        return ndsEntities;
    }

    public void setNdsEntities(List<NdsEntity> ndsEntities) {
        this.ndsEntities = ndsEntities;
    }

    public List<MvUlFounderFl> getMvUlFounderFls() {
        return mvUlFounderFls;
    }

    public void setMvUlFounderFls(List<MvUlFounderFl> mvUlFounderFls) {
        this.mvUlFounderFls = mvUlFounderFls;
    }

    public List<BlockEsf> getBlockEsfs() {
        return blockEsfs;
    }
    public void setBlockEsfs(List<BlockEsf> blockEsfs) {
        this.blockEsfs = blockEsfs;
    }

    public List<DrugAddicts> getDrugAddicts() {
        return drugAddicts;
    }

    public void setDrugAddicts(List<DrugAddicts> drugAddicts) {
        this.drugAddicts = drugAddicts;
    }

    private List<FirstCreditBureauEntity> FirstCreditBureauEntities;  //risk
    private boolean Person_with_risk;

    public boolean isPerson_with_risk() {
        return Person_with_risk;
    }

    public void setPerson_with_risk(boolean person_with_risk) {
        Person_with_risk = person_with_risk;
    }

    public List<MzEntity> getMzEntities() {
        return mzEntities;
    }

    public List<BeneficiariesList> getBeneficiariesLists() {
        return beneficiariesLists;
    }

    public void setBeneficiariesLists(List<BeneficiariesList> beneficiariesLists) {
        this.beneficiariesLists = beneficiariesLists;
    }

    public void setMzEntities(List<MzEntity> mzEntities) {
        this.mzEntities = mzEntities;
    }

    public List<AdvocateListEntity> getAdvocateListEntities() {
        return advocateListEntities;
    }

    public void setAdvocateListEntities(List<AdvocateListEntity> advocateListEntities) {
        this.advocateListEntities = advocateListEntities;
    }

    public List<ImmoralLifestyle> getAmoral() {
        return amoral;
    }

    public void setAmoral(List<ImmoralLifestyle> amoral) {
        this.amoral = amoral;
    }

    public List<Dismissal> getDismissals() {
        return dismissals;
    }

    public void setDismissals(List<Dismissal> dismissals) {
        this.dismissals = dismissals;
    }

    public List<AuditorsListEntity> getAuditorsListEntities() {
        return auditorsListEntities;
    }

    public void setAuditorsListEntities(List<AuditorsListEntity> auditorsListEntities) {
        this.auditorsListEntities = auditorsListEntities;
    }

    public List<BailiffListEntity> getBailiffListEntities() {
        return bailiffListEntities;
    }

    public void setBailiffListEntities(List<BailiffListEntity> bailiffListEntities) {
        this.bailiffListEntities = bailiffListEntities;
    }

    public List<AccountantListEntity> getAccountantListEntities() {
        return accountantListEntities;
    }

    public void setAccountantListEntities(List<AccountantListEntity> accountantListEntities) {
        this.accountantListEntities = accountantListEntities;
    }

    public List<TIpEntity> getTIpEntity() {
        return TIpEntity;
    }

    public void setTIpEntity(List<TIpEntity> TIpEntity) {
        this.TIpEntity = TIpEntity;
    }

    public List<IpgoEmailEntity> getIpgoEmailEntities() {
        return ipgoEmailEntities;
    }

    public void setIpgoEmailEntities(List<IpgoEmailEntity> ipgoEmailEntities) {
        this.ipgoEmailEntities = ipgoEmailEntities;
    }

    public List<FirstCreditBureauEntity> getFirstCreditBureauEntities() {
        return FirstCreditBureauEntities;
    }

    public void setFirstCreditBureauEntities(List<FirstCreditBureauEntity> firstCreditBureauEntities) {
        FirstCreditBureauEntities = firstCreditBureauEntities;
    }

    public List<Msh> getMshes() {
        return mshes;
    }

    public void setMshes(List<Msh> mshes) {
        this.mshes = mshes;
    }

    public List<MilitaryAccounting2Entity> getMilitaryAccounting2Entities() {
        return militaryAccounting2Entities;
    }

    public void setMilitaryAccounting2Entities(List<MilitaryAccounting2Entity> militaryAccounting2Entities) {
        this.militaryAccounting2Entities = militaryAccounting2Entities;
    }

    public List<FlPensionFinal> getFlPensionContrs() {
        return flPensionContrs;
    }

    public void setFlPensionContrs(List<FlPensionFinal> flPensionContrs) {
        this.flPensionContrs = flPensionContrs;
    }


    public List<MvRnOld> getMvRnOlds() {
        return mvRnOlds;
    }

    public void setMvRnOlds(List<MvRnOld> mvRnOlds) {
        this.mvRnOlds = mvRnOlds;
    }



    public List<ConvictsTerminatedByRehab> getConvictsTerminatedByRehabs() {
        return convictsTerminatedByRehabs;
    }

    public void setConvictsTerminatedByRehabs(List<ConvictsTerminatedByRehab> convictsTerminatedByRehabs) {
        this.convictsTerminatedByRehabs = convictsTerminatedByRehabs;
    }

    public List<ConvictsJustified> getConvictsJustifieds() {
        return convictsJustifieds;
    }

    public void setConvictsJustifieds(List<ConvictsJustified> convictsJustifieds) {
        this.convictsJustifieds = convictsJustifieds;
    }

    public List<Bankrot> getBankrots() {
        return bankrots;
    }

    public void setBankrots(List<Bankrot> bankrots) {
        this.bankrots = bankrots;
    }

    public List<MillitaryAccount> getMillitaryAccounts() {
        return millitaryAccounts;
    }

    public void setMillitaryAccounts(List<MillitaryAccount> millitaryAccounts) {
        this.millitaryAccounts = millitaryAccounts;
    }

    public List<FlContacts> getContacts() {
        return contacts;
    }

    public void setContacts(List<FlContacts> contacts) {
        this.contacts = contacts;
    }

    public List<Criminals> getCriminals() {
        return criminals;
    }

    public void setCriminals(List<Criminals> criminals) {
        this.criminals = criminals;
    }

    public List<School> getSchools() {
        return schools;
    }

    public void setSchools(List<School> schools) {
        this.schools = schools;
    }

    public List<Universities> getUniversities() {
        return universities;
    }

    public void setUniversities(List<Universities> universities) {
        this.universities = universities;
    }

    public List<MvIinDoc> getMvIinDocs() {
        return mvIinDocs;
    }

    public void setMvIinDocs(List<MvIinDoc> mvIinDocs) {
        this.mvIinDocs = mvIinDocs;
    }

    public List<Pdl> getPdls() {
        return pdls;
    }

    public void setPdls(List<Pdl> pdls) {
        this.pdls = pdls;
    }

    public List<RegAddressFl> getRegAddressFls() {
        return regAddressFls;
    }

    public void setRegAddressFls(List<RegAddressFl> regAddressFls) {
        this.regAddressFls = regAddressFls;
    }

    public List<FlRelativiesDTO> getFl_relatives() {
        return fl_relatives;
    }

    public void setFl_relatives(List<FlRelativiesDTO> fl_relatives) {
        this.fl_relatives = fl_relatives;
    }

    public List<Orphans> getOrphanss() {
        return orphanss;
    }

    public void setOrphanss(List<Orphans> orphanss) {
        this.orphanss = orphanss;
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

    public List<Orphans> getOrphans() {
        return orphanss;
    }

    public void setOrphans(List<Orphans> orphans) {
        this.orphanss = orphans;
    }

    public List<Omn> getOmns() {
        return omns;
    }

    public void setOmns(List<Omn> omns) {
        this.omns = omns;
    }

    public List<MvFl> getMvFls() {
        return mvFls;
    }

    public void setMvFls(List<MvFl> mvFls) {
        this.mvFls = mvFls;
    }

    public List<MvAutoFl> getMvAutoFls() {
        return mvAutoFls;
    }

    public void setMvAutoFls(List<MvAutoFl> mvAutoFls) {
        this.mvAutoFls = mvAutoFls;
    }

    public List<PhotoDb> getPhotoDbf() {
        return photoDbf;
    }

    public void setPhotoDbf(List<PhotoDb> photoDbf) {
        this.photoDbf = photoDbf;
    }
}
