package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.Dismissal;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface DismissalRepo extends JpaRepository<Dismissal, Integer> {
    @Query(value= "select * from imp_risk.dismissal where iin = ?1", nativeQuery = true)
    List<Dismissal> getDismissalByIIN(String IIN);
}
