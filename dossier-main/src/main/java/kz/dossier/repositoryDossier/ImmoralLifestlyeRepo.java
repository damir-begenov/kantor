package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.ImmoralLifestyle;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ImmoralLifestlyeRepo extends JpaRepository<ImmoralLifestyle, Integer> {
    @Query(value= "select * from imp_kfm_fl.immoral_lifestyle where iin = ?1", nativeQuery = true)
    List<ImmoralLifestyle> getImmoByIIN(String IIN);
}
