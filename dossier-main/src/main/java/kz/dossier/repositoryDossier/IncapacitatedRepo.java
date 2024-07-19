package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.Incapacitated;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface IncapacitatedRepo extends JpaRepository<Incapacitated, Integer> {
    @Query(value = "select * from imp_risk.incapacitated where iin = ?1",nativeQuery = true)
    List<Incapacitated> getIncapacitatedByIIN(String iin);
}
