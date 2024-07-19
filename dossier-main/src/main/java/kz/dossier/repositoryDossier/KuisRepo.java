package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.Kuis;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface KuisRepo extends JpaRepository<Kuis, Integer> {
    @Query(value = "select * from imp_risk.kuis where iin = ?1", nativeQuery = true)
    List<Kuis> getKuisByIIN(String iin);
}
