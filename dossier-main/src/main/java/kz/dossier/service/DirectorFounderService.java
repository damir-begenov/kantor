package kz.dossier.service;

import kz.dossier.modelsDossier.FlFounderLeaderDTO;
import kz.dossier.modelsDossier.MvUlFounderFl;
import kz.dossier.modelsDossier.MvUlFounderUl;
import kz.dossier.modelsDossier.MvUlLeaderEntity;
import kz.dossier.repositoryDossier.MvUlFounderFlRepo;
import kz.dossier.repositoryDossier.MvUlFounderUlRepo;
import kz.dossier.repositoryDossier.MvUlLeaderEntityRepo;
import lombok.AllArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.Collections;
import java.util.List;
import java.util.Optional;

@Service
@AllArgsConstructor
public class DirectorFounderService {
    private final MvUlFounderFlRepo mvUlFounderFlRepo;
    private final MvUlLeaderEntityRepo mvUlLeaderEntityRepo;

    public FlFounderLeaderDTO getDirectorOrFounder(String iin){
        FlFounderLeaderDTO flFounderLeaderDTO = new FlFounderLeaderDTO();
        try {
            List<MvUlFounderFl> mvUlFounderFlList = mvUlFounderFlRepo.getUsersByLikeIIN(iin);
            flFounderLeaderDTO.setMvUlFounderFlList(mvUlFounderFlList);
        }catch (Exception e){
            System.out.println("mvUlFounderFlList error");
        }
        try {
            List<MvUlLeaderEntity> mvUlLeaderEntities = mvUlLeaderEntityRepo.getUsersByLikeIin(iin);
            flFounderLeaderDTO.setMvUlLeaderEntities(mvUlLeaderEntities);
        }catch (Exception e){
            System.out.println("mvUlFounderFlList error");
        }

        int sum = (Optional.ofNullable(flFounderLeaderDTO.getMvUlFounderFlList()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0) +
                (Optional.ofNullable(flFounderLeaderDTO.getMvUlLeaderEntities()).orElse(Collections.emptyList()).size() > 0 ? 1 : 0);
        flFounderLeaderDTO.setQuantity(sum);
        return flFounderLeaderDTO;
    }
}
