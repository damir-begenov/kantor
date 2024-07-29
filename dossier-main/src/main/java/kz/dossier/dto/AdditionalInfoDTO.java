package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.*;
import kz.dossier.modelsRisk.Pdl;

public class AdditionalInfoDTO {
    private List<Universities> universities;
    private List<School> schools;

    private List<MvRnOld> mvRnOlds; //mvRn
    private List<CommodityProducer> commodityProducers;

    private List<MvAutoFl> mvAutoFls;
    private List<Equipment> equipment;
    private List<MilitaryAccountingDTO> militaryAccounting2Entities;

    private List<MvUlLeader> ul_leaderList; //Сведения об участии в ЮЛ
    private List<FlPensionFinal> flPensionContrs;


    public void setNumber(int number) {
        this.number = number;
    }

    private int number;

    public int getNumber() {
        return number;
    }

    public void setNumber() {
        int s = 0;

        // Check if lists are not null and not empty
        if (this.flPensionContrs != null && !this.flPensionContrs.isEmpty()) {
            ++s;
        }
        if (this.equipment != null && !this.equipment.isEmpty()) {
            ++s;
        }
        if (this.universities != null && !this.universities.isEmpty()) {
            ++s;
        }
        if (this.schools != null && !this.schools.isEmpty()) {
            ++s;
        }
        if (this.mvRnOlds != null && !this.mvRnOlds.isEmpty()) {
            ++s;
        }
        if (this.mvAutoFls != null && !this.mvAutoFls.isEmpty()) {
            ++s;
        }
        if (this.militaryAccounting2Entities != null && !this.militaryAccounting2Entities.isEmpty()) {
            ++s;
        }
        if (this.ul_leaderList != null && !this.ul_leaderList.isEmpty()) {
            ++s;
        }
        if (this.commodityProducers != null && !this.commodityProducers.isEmpty()) {
            ++s;
        }

        this.number = ++s;
    }

    List<PensionListDTO> pensions;

    public List<CommodityProducer> getCommodityProducers() {
        return commodityProducers;
    }

    public void setCommodityProducers(List<CommodityProducer> commodityProducers) {
        this.commodityProducers = commodityProducers;
    }

    public void setNumber(int number) {
        this.number = number;
    }

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
    public void setMilitaryAccounting2Entities(List<MilitaryAccountingDTO> militaryAccounting2Entities) {
        this.militaryAccounting2Entities = militaryAccounting2Entities;
    }

    public List<MilitaryAccountingDTO> getMilitaryAccounting2Entities() {
        return militaryAccounting2Entities;
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
