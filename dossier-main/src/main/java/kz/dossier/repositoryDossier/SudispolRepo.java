package kz.dossier.repositoryDossier;

import kz.dossier.modelsRisk.Sudispol;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface SudispolRepo extends JpaRepository<Sudispol, Long> {
    @Query(value = "select * from imp_risk.sudispol where debtor_iin = ?1", nativeQuery = true)
    List<Sudispol> findAllbyIin(String iin);
}
