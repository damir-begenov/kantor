package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.Unemployed;
import kz.dossier.modelsDossier.Universities;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface UnemployedRepo extends JpaRepository<Unemployed, Long> {
    @Query(value = "select * from imp_risk.unemployed where iin = ?1", nativeQuery = true)
    List<Unemployed> getUnemployedByIIN(String iin);
}
