package kz.dossier.service;

import kz.dossier.dto.ULAdditionalInfoDTO;
import kz.dossier.dto.ULULMemberDTO;
import kz.dossier.modelsDossier.*;
import kz.dossier.modelsRisk.FpgTempEntity;
import kz.dossier.repositoryDossier.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
@Service
public class ULAdditionalService {
    @Autowired
    private MvRnOldRepo mv_rn_oldRepo;
    @Autowired
    private MvUlRepo mv_ul_repo;
    @Autowired
    private MvAutoFlRepo mvAutoFlRepo;
    @Autowired
    private EquipmentRepo equipment_repo;
    @Autowired
    private TrainsRepo trainsRepo;
    @Autowired
    private WaterTransportRepo waterTransportRepo;
    @Autowired
    private AutoTransportRepo autoTransportRepo;
    @Autowired
    private AutoPostanovkaRepo autoPostanovkaRepo;
    @Autowired
    private AutoSnyatieRepo autoSnyatieRepo;
    @Autowired
    private AviaTransportRepo aviaTransportRepo;
    @Autowired
    FpgTempEntityRepo fpgTempEntityRepo;
    @Autowired
    private MvUlLeaderRepository mvUlLeaderRepository;
    @Autowired
    MvUlFounderUlRepo mvUlFounderUlRepo;
    @Autowired
    FlRiskServiceImpl flRislimplementation;
    public ULAdditionalInfoDTO additionalByBin(String bin) {
        ULAdditionalInfoDTO result = new ULAdditionalInfoDTO();
        try {
            List<MvRnOld> mvRnOlds = mv_rn_oldRepo.getUsersByLike(bin);
            List<MvRnOld> list = setNamesByBin(mvRnOlds);
            result.setMvRnOlds(list);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<MvAutoFl> myMv_auto_fl =  mvAutoFlRepo.getUsersByLike(bin);
            try {
                result.setMvAutoFls(myMv_auto_fl);
            } catch (Exception e) {
                System.out.println("mv_auto_fl Error: " + e);
            }

        } catch (Exception e){
            System.out.println("mv_auto_fl WRAP Error:" + e);
        }
        try {
            List<Equipment> myEquipment =  equipment_repo.getUsersByLike(bin);
            result.setEquipment(myEquipment);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<Trains> trains =  trainsRepo.getByIIN(bin);
            result.setTrains(trains);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<WaterTransport> waterTransports =  waterTransportRepo.getWaterByIin(bin);
            result.setWaterTransports(waterTransports);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<AutoTransport> autoTransports =  autoTransportRepo.getAutoByIin(bin);
            result.setAutoTransports(autoTransports);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            result.setAutoPostanovkas(autoPostanovkaRepo.getAutoPostanovkaByBin(bin));
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            result.setAutoSnyaties(autoSnyatieRepo.getAutoSnyatieByIin(bin));
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<AviaTransport> aviaTransports =  aviaTransportRepo.getAviaByIin(bin);
            result.setAviaTransports(aviaTransports);
        } catch (Exception e){
            System.out.println("Error:" + e);
        }
        try {
            List<FpgTempEntity> fpgTempEntities = fpgTempEntityRepo.getUsersByLike(bin);
            List<String> fpgs = new ArrayList<>();
            fpgTempEntities.forEach(x -> {
                fpgs.add(x.getBeneficiary() != null ? x.getBeneficiary() : "");
            });
            result.setFpg(fpgs);
        } catch (Exception e) {
            System.out.println("Error: " + e);
        }
        List<ULULMemberDTO> ululMemberDTOS = new ArrayList<>();
        try {
            List<MvUlFounderUl> ulFounders = mvUlFounderUlRepo.getUsersByLike(bin);
            ulFounders.forEach(x -> {
                ULULMemberDTO obj = new ULULMemberDTO();
                obj.setDate(x.getRegDate());
                obj.setBinIin(x.getFounderBin());
                obj.setName(x.getFounderNameRu());
                if (x.isCurrent()) {
                    obj.setPosition("Учредитель ЮЛ");
                } else {
                    obj.setPosition("Учредитель ЮЛ (Исторический)");
                }
//                obj.setRisksNumber();
                ululMemberDTOS.add(obj);
            });
        } catch (Exception e) {
            System.out.println(e);
        }
        try {
            List<MvUlLeader> flFounders = mvUlLeaderRepository.findAllByBinOrg(bin);
            flFounders.forEach(x -> {
                ULULMemberDTO obj = new ULULMemberDTO();
                obj.setDate(x.getReg_date() != null ? x.getReg_date().toString() : "");
                obj.setBinIin(x.getIin());
                try {
                    String name = x.getLast_name() + " " + x.getFirst_name() + x.getPatronymic();
                    obj.setName(name);
                } catch (Exception e) {

                }
                if (x.isIs_curr()) {
                    obj.setPosition("Директор");
                } else {
                    obj.setPosition("Директор (Исторический)");
                }
                FLRiskDto risks = flRislimplementation.findFlRiskByIin(x.getIin());
                obj.setRisksNumber(risks.getQuantity());
                ululMemberDTOS.add(obj);

            });
        } catch (Exception e) {
            System.out.println(e);
        }
        result.setUlMembers(ululMemberDTOS);

        return result;
    }

    private List<MvRnOld> setNamesByBin(List<MvRnOld> list) {
        for (MvRnOld a : list) {
            String name = mv_ul_repo.getNameByBin(a.getOwner_iin_bin());
            if (name != null) {
                a.setOwner_full_name(name);
            }
        }
        return list;
    }
}
