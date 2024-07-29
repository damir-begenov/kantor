package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.MvUlLeaderEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface MvUlLeaderEntityRepo extends JpaRepository<MvUlLeaderEntity, Long> {
    @Query(value= "select * from imp_kfm_ul.ul_leader mv_ul0_ where mv_ul0_.bin_org = ?1 ", nativeQuery = true)
    List<MvUlLeaderEntity> getUsersByLike(String iin);
    @Query(value= "select * from imp_kfm_ul.ul_leader mv_ul0_ where mv_ul0_.iin = ?1 ", nativeQuery = true)
    List<MvUlLeaderEntity> getUsersByLikeIin(String iin);
}