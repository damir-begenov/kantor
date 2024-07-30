package kz.dossier.dto;

import kz.dossier.modelsDossier.*;

import java.util.List;

public class ULAdditionalInfoDTO {
    private List<MvRnOld> mvRnOlds; //mvRn
    private List<MvAutoFl> mvAutoFls;
    private List<AutoTransport> autoTransports; // управление
    private List<AviaTransport> aviaTransports;
    private List<WaterTransport> waterTransports;
    private List<AutoPostanovka> autoPostanovkas;
    private List<AutoSnyatie> autoSnyaties;
    private List<Trains> trains;
    private List<Equipment> equipment;

    private List<ULULMemberDTO> ulMembers;

    private List<String> fpg;

    public List<MvRnOld> getMvRnOlds() {
        return mvRnOlds;
    }

    public void setMvRnOlds(List<MvRnOld> mvRnOlds) {
        this.mvRnOlds = mvRnOlds;
    }

    public List<MvAutoFl> getMvAutoFls() {
        return mvAutoFls;
    }

    public void setMvAutoFls(List<MvAutoFl> mvAutoFls) {
        this.mvAutoFls = mvAutoFls;
    }

    public List<AutoTransport> getAutoTransports() {
        return autoTransports;
    }

    public void setAutoTransports(List<AutoTransport> autoTransports) {
        this.autoTransports = autoTransports;
    }

    public List<AviaTransport> getAviaTransports() {
        return aviaTransports;
    }

    public void setAviaTransports(List<AviaTransport> aviaTransports) {
        this.aviaTransports = aviaTransports;
    }

    public List<WaterTransport> getWaterTransports() {
        return waterTransports;
    }

    public void setWaterTransports(List<WaterTransport> waterTransports) {
        this.waterTransports = waterTransports;
    }

    public List<AutoPostanovka> getAutoPostanovkas() {
        return autoPostanovkas;
    }

    public void setAutoPostanovkas(List<AutoPostanovka> autoPostanovkas) {
        this.autoPostanovkas = autoPostanovkas;
    }

    public List<AutoSnyatie> getAutoSnyaties() {
        return autoSnyaties;
    }

    public void setAutoSnyaties(List<AutoSnyatie> autoSnyaties) {
        this.autoSnyaties = autoSnyaties;
    }

    public List<Trains> getTrains() {
        return trains;
    }

    public void setTrains(List<Trains> trains) {
        this.trains = trains;
    }

    public List<Equipment> getEquipment() {
        return equipment;
    }

    public void setEquipment(List<Equipment> equipment) {
        this.equipment = equipment;
    }

    public List<ULULMemberDTO> getUlMembers() {
        return ulMembers;
    }

    public void setUlMembers(List<ULULMemberDTO> ulMembers) {
        this.ulMembers = ulMembers;
    }

    public List<String> getFpg() {
        return fpg;
    }

    public void setFpg(List<String> fpg) {
        this.fpg = fpg;
    }
}
