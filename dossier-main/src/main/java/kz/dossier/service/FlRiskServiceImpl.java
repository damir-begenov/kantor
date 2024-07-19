package kz.dossier.service;

import kz.dossier.modelsDossier.FLRiskDto;
import kz.dossier.repositoryDossier.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.Collections;
import java.util.Optional;

@Service
public class FlRiskServiceImpl {

    @Autowired
    private AdmRepo admRepo;
    @Autowired
    private OrphansRepo orphans_repo;
    @Autowired
    private OmnRepo omn_repo;
    @Autowired
    private BankrotRepo bankrotRepo;
    @Autowired
    private BlockEsfRepo block_esfRepo;
    @Autowired
    private ConvictsJustifiedRepo convicts_justifiedRepo;
    @Autowired
    private ConvictsTerminatedtByRehabRepo convicts_terminated_by_rehabRepo;
    @Autowired
    private CriminalsRepo criminalsRepo;
    @Autowired
    private DismissalRepo dismissalRepo;
    @Autowired
    private FirstCreditBureauEntityRepo firstCreditBureauEntityRepo;
    @Autowired
    private FpgTempEntityRepo fpgTempEntityRepo;
    @Autowired
    private MzEntityRepo mzEntityRepo;
    @Autowired
    private NdsEntityRepo ndsEntityRepo;
    @Autowired
    private OpgRepo opgRepo;
    @Autowired
    private WantedListRepo wantedListRepo;
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
    public FLRiskDto findFlRiskByIin(String iin){
        FLRiskDto flRiskDto = new FLRiskDto();

        flRiskDto.setIin(iin);
        flRiskDto.setBankrots(bankrotRepo.getbankrotByByIIN(iin));
        flRiskDto.setIncapacitateds(incapacitatedRepo.getIncapacitatedByIIN(iin));
        flRiskDto.setKuis(kuisRepo.getKuisByIIN(iin));
        flRiskDto.setAdms(admRepo.getUsersByLike(iin));
        flRiskDto.setCriminals(criminalsRepo.getcriminalsByByIIN(iin));
        flRiskDto.setDismissals(dismissalRepo.getDismissalByIIN(iin));
        flRiskDto.setBeneficiariesLists(beneficiariesListRepo.getBenByIIN(iin));
        flRiskDto.setDormants(dormantRepo.getUsersByLike(iin));
        flRiskDto.setBlockEsfs(block_esfRepo.getblock_esfByIIN(iin));
        flRiskDto.setMshes(mshRepo.getUsersByLike(iin));
        flRiskDto.setConvictsAbroads(convictsAbroadRepo.getConvictsAbroadByIIN(iin));
        flRiskDto.setConvictsJustifieds(convicts_justifiedRepo.getconvicts_justifiedByByIIN(iin));
        flRiskDto.setDrugAddicts(drugAddictsRepo.getDrugAddictsByIIN(iin));
        flRiskDto.setOmns(omn_repo.getUsersByLike(iin));
        flRiskDto.setOrphans(orphans_repo.getUsersByLike(iin));
        flRiskDto.setMzEntities(mzEntityRepo.getopgByIIN(iin));
        flRiskDto.setWantedListEntities(wantedListRepo.getByIIN(iin));
        flRiskDto.setFirstCreditBureauEntities(firstCreditBureauEntityRepo.getUsersByLike(iin));
        flRiskDto.setConvictsTerminatedByRehabs(convicts_terminated_by_rehabRepo.getconvicts_terminated_by_rehabByByIIN(iin));

        int sum = (Optional.ofNullable(flRiskDto.getAdms()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getBankrots()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getDormants()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                Optional.ofNullable(flRiskDto.getFirstCreditBureauEntities()).orElse(Collections.emptyList()).size() +
                (Optional.ofNullable(flRiskDto.getAccountantListEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getConvictsAbroads()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getIncapacitateds()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getKuis()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getDismissals()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getMshes()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getDrugAddicts()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getWantedListEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getFpgTempEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getNdsEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flRiskDto.getConvictsTerminatedByRehabs()).orElse(Collections.emptyList()).isEmpty() ? 0 : 1);

        flRiskDto.setQuantity(sum);

        return flRiskDto;
    }

    //--Осужденные
}
