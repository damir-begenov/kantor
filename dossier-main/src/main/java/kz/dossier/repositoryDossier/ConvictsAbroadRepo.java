package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.ConvictsAbroad;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface ConvictsAbroadRepo extends JpaRepository<ConvictsAbroad, Integer> {
    @Query(value = "select * from imp_risk.convicts_abroad where iin = ?1", nativeQuery = true)
    List<ConvictsAbroad> getConvictsAbroadByIIN(String IIN);
}
