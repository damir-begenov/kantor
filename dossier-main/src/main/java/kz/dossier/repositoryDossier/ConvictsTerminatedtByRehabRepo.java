package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.ConvictsTerminatedByRehab;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface ConvictsTerminatedtByRehabRepo extends JpaRepository<ConvictsTerminatedByRehab, Long> {
    @Query(value= "select * from imp_risk.convicts_terminated_by_rehab where iin = ?1", nativeQuery = true)
    List<ConvictsTerminatedByRehab> getconvicts_terminated_by_rehabByByIIN(String IIN);
}