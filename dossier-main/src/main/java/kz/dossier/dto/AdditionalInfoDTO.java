package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.Equipment;
import kz.dossier.modelsDossier.FlPensionFinal;
import kz.dossier.modelsDossier.MilitaryAccounting2Entity;
import kz.dossier.modelsDossier.MillitaryAccount;
import kz.dossier.modelsDossier.MvAutoFl;
import kz.dossier.modelsDossier.MvRnOld;
import kz.dossier.modelsDossier.MvUlFounderFl;
import kz.dossier.modelsDossier.MvUlLeader;
import kz.dossier.modelsDossier.School;
import kz.dossier.modelsDossier.Universities;

public class AdditionalInfoDTO {
    private List<Universities> universities;
    private List<School> schools;

    private List<MvRnOld> mvRnOlds; //mvRn

    private List<MvAutoFl> mvAutoFls;
    private List<Equipment> equipment;
    private List<MillitaryAccount> millitaryAccounts;
    private List<MilitaryAccounting2Entity> militaryAccounting2Entities;

    private List<MvUlLeader> ul_leaderList; //Сведения об участии в ЮЛ
    private List<FlPensionFinal> flPensionContrs;

    List<PensionListDTO> pensions;

    public List<PensionListDTO> getPensions() {
        return pensions;
    }
    public void setPensions(List<PensionListDTO> pensions) {
        this.pensions = pensions;
    }

    public List<Equipment> getEquipment() {
        return equipment;
    }
    public List<FlPensionFinal> getFlPensionContrs() {
        return flPensionContrs;
    }
    public List<MilitaryAccounting2Entity> getMilitaryAccounting2Entities() {
        return militaryAccounting2Entities;
    }
    public List<MillitaryAccount> getMillitaryAccounts() {
        return millitaryAccounts;
    }
    public List<MvAutoFl> getMvAutoFls() {
        return mvAutoFls;
    }
    public List<MvRnOld> getMvRnOlds() {
        return mvRnOlds;
    }
    public List<School> getSchools() {
        return schools;
    }
    public List<MvUlLeader> getUl_leaderList() {
        return ul_leaderList;
    }
    public List<Universities> getUniversities() {
        return universities;
    }
    public void setEquipment(List<Equipment> equipment) {
        this.equipment = equipment;
    }
    public void setFlPensionContrs(List<FlPensionFinal> flPensionContrs) {
        this.flPensionContrs = flPensionContrs;
    }
    public void setMilitaryAccounting2Entities(List<MilitaryAccounting2Entity> militaryAccounting2Entities) {
        this.militaryAccounting2Entities = militaryAccounting2Entities;
    }
    public void setMillitaryAccounts(List<MillitaryAccount> millitaryAccounts) {
        this.millitaryAccounts = millitaryAccounts;
    }
    public void setMvAutoFls(List<MvAutoFl> mvAutoFls) {
        this.mvAutoFls = mvAutoFls;
    }
    public void setMvRnOlds(List<MvRnOld> mvRnOlds) {
        this.mvRnOlds = mvRnOlds;
    }
    public void setSchools(List<School> schools) {
        this.schools = schools;
    }
    public void setUl_leaderList(List<MvUlLeader> ul_leaderList) {
        this.ul_leaderList = ul_leaderList;
    }
    public void setUniversities(List<Universities> universities) {
        this.universities = universities;
    }
    

}
