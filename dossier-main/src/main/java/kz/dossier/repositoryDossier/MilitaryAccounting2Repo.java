package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.MilitaryAccounting2Entity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository

public interface MilitaryAccounting2Repo extends JpaRepository<MilitaryAccounting2Entity, Long> {
    @Query(value= "select * from imp_kfm_fl.military_accounting where iin = ?1 ", nativeQuery = true)
    List<MilitaryAccounting2Entity> getUsersByLike(String iin);
}