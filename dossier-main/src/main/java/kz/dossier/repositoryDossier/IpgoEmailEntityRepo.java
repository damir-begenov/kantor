package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.IpgoEmailEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface IpgoEmailEntityRepo extends JpaRepository<IpgoEmailEntity, Long> {
    @Query(value= "select * from imp_kfm_fl.ipgo_email where iin = ?1 ", nativeQuery = true)
    List<IpgoEmailEntity> getUsersByLike(String iin);
}