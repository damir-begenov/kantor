package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.*;
import kz.dossier.modelsDossier.AccountantListEntity;
import kz.dossier.modelsRisk.Pdl;

public class GeneralInfoDTO {
    private Double percent;
    private List<FlContacts> contacts;

    public Double getPercent() {
        return percent;
    }

    public void setPercent(Double percent) {
        this.percent = percent;
    }

    private List<SearchResultModelFL> sameAddressFls;
    private List<IndividualEntrepreneur> individualEntrepreneurs;
    private List<KX> kxes;
    private List<AccountantListEntity> accountantListEntities; // БУХГАЛТЭРЫ
    private List<AdvocateListEntity> advocateListEntities;
    private List<AuditorsListEntity> auditorsListEntities;
    private List<BailiffListEntity> bailiffListEntities; //частный суд исполнитель
    private List<IpgoEmailEntity> ipgoEmailEntities;
    private List<MvUlFounderFl> mvUlFounderFls;
    private List<MvUlLeaderEntity> ul_leaderList; //Сведения об участии в ЮЛ

    public void setUl_leaderList(List<MvUlLeaderEntity> ul_leaderList) {
        this.ul_leaderList = ul_leaderList;
    }

    public List<MvUlLeaderEntity> getUl_leaderList() {
        return ul_leaderList;
    }

    public List<MvUlFounderFl> getMvUlFounderFls() {
        return mvUlFounderFls;
    }


    public void setMvUlFounderFls(List<MvUlFounderFl> mvUlFounderFls) {
        this.mvUlFounderFls = mvUlFounderFls;
    }


    public List<IpgoEmailEntity> getIpgoEmailEntities() {
        return ipgoEmailEntities;
    }

    public void setIpgoEmailEntities(List<IpgoEmailEntity> ipgoEmailEntities) {
        this.ipgoEmailEntities = ipgoEmailEntities;
    }

    public List<AdvocateListEntity> getAdvocateListEntities() {
        return advocateListEntities;
    }

    public void setAdvocateListEntities(List<AdvocateListEntity> advocateListEntities) {
        this.advocateListEntities = advocateListEntities;
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

    private List<Lawyers> lawyers;
    private ChangeFio changeFio;

    private List<Pdl> pdls;

    public List<KX> getKxes() {
        return kxes;
    }

    public void setKxes(List<KX> kxes) {
        this.kxes = kxes;
    }

    public List<AccountantListEntity> getAccountantListEntities() {
        return accountantListEntities;
    }

    public void setAccountantListEntities(List<AccountantListEntity> accountantListEntities) {
        this.accountantListEntities = accountantListEntities;
    }

    public List<Pdl> getPdls() {
        return pdls;
    }

    public void setPdls(List<Pdl> pdls) {
        this.pdls = pdls;
    }

    public ChangeFio getChangeFio() {
        return changeFio;
    }


    public List<IndividualEntrepreneur> getIndividualEntrepreneurs() {
        return individualEntrepreneurs;
    }

    public void setIndividualEntrepreneurs(List<IndividualEntrepreneur> individualEntrepreneurs) {
        this.individualEntrepreneurs = individualEntrepreneurs;
    }
    public void setChangeFio(ChangeFio changeFio) {
        this.changeFio = changeFio;
    }

    public List<Lawyers> getLawyers() {
        return lawyers;
    }

    public void setLawyers(List<Lawyers> lawyers) {
        this.lawyers = lawyers;
    }

    public List<FlContacts> getContacts() {
        return contacts;
    }
    public List<SearchResultModelFL> getSameAddressFls() {
        return sameAddressFls;
    }


    public void setContacts(List<FlContacts> contacts) {
        this.contacts = contacts;
    }
    public void setSameAddressFls(List<SearchResultModelFL> sameAddressFls) {
        this.sameAddressFls = sameAddressFls;
    }
}
