package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.ChangeFio;
import kz.dossier.modelsDossier.FlContacts;
import kz.dossier.modelsDossier.Lawyers;
import kz.dossier.modelsDossier.IndividualEntrepreneur;
import kz.dossier.modelsDossier.SearchResultModelFL;
import kz.dossier.modelsRisk.Pdl;

public class GeneralInfoDTO {
    private List<FlContacts> contacts;
    private List<SearchResultModelFL> sameAddressFls;
    private List<IndividualEntrepreneur> individualEntrepreneurs;

    private List<Lawyers> lawyers;
    private ChangeFio changeFio;

    private List<Pdl> pdls;

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
