package kz.dossier.service;

import kz.dossier.dto.UlRiskDTO;
import kz.dossier.repositoryDossier.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.Collections;
import java.util.Optional;

@Service
public class UlRiskServiceImpl {
    @Autowired
    private OmnRepo omn_repo;
    @Autowired
    private BankrotRepo bankrotRepo;
    @Autowired
    private BlockEsfRepo block_esfRepo;
    @Autowired
    private CriminalsRepo criminalsRepo;
    @Autowired
    private FirstCreditBureauEntityRepo firstCreditBureauEntityRepo;
    @Autowired
    private FpgTempEntityRepo fpgTempEntityRepo;
    @Autowired
    private NdsEntityRepo ndsEntityRepo;
    @Autowired
    private OpgRepo opgRepo;
    @Autowired
    DormantRepo dormantRepo;
    @Autowired
    MshRepo mshRepo;
    @Autowired
    BeneficiariesListRepo beneficiariesListRepo;
    @Autowired
    ConvictsAbroadRepo convictsAbroadRepo;
    @Autowired
    KuisRepo kuisRepo;
    @Autowired
    IncapacitatedRepo incapacitatedRepo;
    @Autowired
    DrugAddictsRepo drugAddictsRepo;
    @Autowired
    ImmoralLifestlyeRepo immoralLifestlyeRepo;
    @Autowired
    VerificationFieldRepo verificationFieldRepo;

    public UlRiskDTO findULRiskByIin(String bin) {
        UlRiskDTO ulRiskDto = new UlRiskDTO();
        ulRiskDto.setBin(bin);

        try {
            ulRiskDto.setBankrots(bankrotRepo.getbankrotByByIIN(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        try {
            ulRiskDto.setBlockEsfs(block_esfRepo.getblock_esfByIIN(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        try {
            ulRiskDto.setCriminals(criminalsRepo.getcriminalsByByIIN(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        try {
            ulRiskDto.setDormants(dormantRepo.getUsersByLike(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        try {
            ulRiskDto.setKuis(kuisRepo.getKuisByBin(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        try {
            ulRiskDto.setFirstCreditBureauEntities(firstCreditBureauEntityRepo.getUsersByLike(bin));
        } catch (Exception e) {
            // Catch block left empty
        }try {
            ulRiskDto.setNdsEntities(ndsEntityRepo.getUsersByLike(bin));
        } catch (Exception e) {
            // Catch block left empty
        }try {
            ulRiskDto.setOmns(omn_repo.getUsersByLikeIin_bin(bin));
        } catch (Exception e) {
            // Catch block left empty
        }try {
            ulRiskDto.setOpgEntities(opgRepo.getopgByIIN(bin));
        } catch (Exception e) {
            // Catch block left empty
        }try {
            ulRiskDto.setFpgTempEntities(fpgTempEntityRepo.getUsersByLike(bin));
        } catch (Exception e) {
            // Catch block left empty
        }try {
            ulRiskDto.setVerificationFields(verificationFieldRepo.getByIIN(bin));
        } catch (Exception e) {
            // Catch block left empty
        }
        int sum =
                (Optional.ofNullable(ulRiskDto.getBankrots()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getDormants()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        Optional.ofNullable(ulRiskDto.getFirstCreditBureauEntities()).orElse(Collections.emptyList()).size() +
                        (Optional.ofNullable(ulRiskDto.getKuis()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getOpgEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getOmns()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getFpgTempEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getCriminals()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getVerificationFields()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getBlockEsfs()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                        (Optional.ofNullable(ulRiskDto.getNdsEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0);

        ulRiskDto.setQuantity(sum);
        return ulRiskDto;
    }

}
