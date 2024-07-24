package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.FlContacts;
import kz.dossier.modelsDossier.SearchResultModelFL;

public class GeneralInfoDTO {
    private List<FlContacts> contacts;
    private List<SearchResultModelFL> sameAddressFls;


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
