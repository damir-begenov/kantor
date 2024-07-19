package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.Msh;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface MshRepo extends JpaRepository<Msh, Long> {
    @Query(value= "select * from imp_kfm_fl.msh where owner_iin_bin = ?1 ", nativeQuery = true)
    List<Msh> getUsersByLike(String iin);
}