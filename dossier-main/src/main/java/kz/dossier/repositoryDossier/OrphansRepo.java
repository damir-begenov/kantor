package kz.dossier.repositoryDossier;

import kz.dossier.modelsRisk.Orphans;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface OrphansRepo extends JpaRepository<Orphans, Long> {
    @Query(value= "select * from imp_risk.orphans orphans0_ where orphans0_.iin = ?1 ", nativeQuery = true)
    List<Orphans> getUsersByLike(String iin);

}
