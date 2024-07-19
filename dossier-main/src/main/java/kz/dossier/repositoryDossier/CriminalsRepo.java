package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.Criminals;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface CriminalsRepo extends JpaRepository<Criminals, Long> {
    @Query(value= "select * from imp_risk.criminals where iin_bin = ?1", nativeQuery = true)
    List<Criminals> getcriminalsByByIIN(String IIN);
}