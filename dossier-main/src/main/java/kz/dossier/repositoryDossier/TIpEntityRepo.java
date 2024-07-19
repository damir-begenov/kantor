package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.TIpEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface TIpEntityRepo extends JpaRepository<TIpEntity, Long> {
    @Query(value= "select * from imp_kfm_fl.mv_ip where iin = ?1 ", nativeQuery = true)
    List<TIpEntity> getUsersByLike(String iin);
}