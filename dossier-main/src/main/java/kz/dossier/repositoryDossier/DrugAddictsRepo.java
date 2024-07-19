package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.DrugAddicts;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository
public interface DrugAddictsRepo extends JpaRepository<DrugAddicts, Integer> {
    @Query(value = "Select * from imp_risk.drug_addicts_pfr where iin = ?1", nativeQuery = true)
    List<DrugAddicts> getDrugAddictsByIIN(String IIN);
}
