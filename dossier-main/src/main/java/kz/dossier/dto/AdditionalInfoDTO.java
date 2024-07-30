package kz.dossier.dto;

import java.util.List;

import kz.dossier.modelsDossier.*;

public class AdditionalInfoDTO {
    private List<Universities> universities;
    private List<School> schools;
    private List<MvUlLeaderEntity> ul_leaderList; //Сведения об участии в ЮЛ

    public List<MvUlLeaderEntity> getUl_leaderList() {
        return ul_leaderList;
    }

    public void setUl_leaderList(List<MvUlLeaderEntity> ul_leaderList) {
        this.ul_leaderList = ul_leaderList;
    }

    private List<KX> kxes;

    public List<KX> getKxes() {
        return kxes;
    }

    public void setKxes(List<KX> kxes) {
        this.kxes = kxes;
    }

    private List<MvRnOld> mvRnOlds; //mvRn
    private List<CommodityProducer> commodityProducers;

    private List<MvAutoFl> mvAutoFls;
    private List<AutoTransport> autoTransports; // управление
    private List<AviaTransport> aviaTransports;
    private List<WaterTransport> waterTransports;
    private List<AutoPostanovka> autoPostanovkas;
    private List<AutoSnyatie> autoSnyaties;

    public List<AutoSnyatie> getAutoSnyaties() {
        return autoSnyaties;
    }

    public void setAutoSnyaties(List<AutoSnyatie> autoSnyaties) {
        this.autoSnyaties = autoSnyaties;
    }

    public List<AutoPostanovka> getAutoPostanovkas() {
        return autoPostanovkas;
    }

    public void setAutoPostanovkas(List<AutoPostanovka> autoPostanovkas) {
        this.autoPostanovkas = autoPostanovkas;
    }

    private List<Trains> trains;
    private List<Equipment> equipment;
    private List<MilitaryAccountingDTO> militaryAccounting2Entities;

    private List<FlPensionFinal> flPensionContrs;



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

    public List<AutoTransport> getAutoTransports() {
        return autoTransports;
    }

    public void setAutoTransports(List<AutoTransport> autoTransports) {
        this.autoTransports = autoTransports;
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
    public List<Universities> getUniversities() {
        return universities;
    }

    public List<Trains> getTrains() {
        return trains;
    }

    public void setTrains(List<Trains> trains) {
        this.trains = trains;
    }

    public List<WaterTransport> getWaterTransports() {
        return waterTransports;
    }

    public void setWaterTransports(List<WaterTransport> waterTransports) {
        this.waterTransports = waterTransports;
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
    public void setUniversities(List<Universities> universities) {
        this.universities = universities;
    }

    public List<AviaTransport> getAviaTransports() {
        return aviaTransports;
    }

    public void setAviaTransports(List<AviaTransport> aviaTransports) {
        this.aviaTransports = aviaTransports;
    }
}
