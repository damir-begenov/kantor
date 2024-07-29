package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.KX;
import kz.dossier.modelsRisk.Dismissal;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface KxRepo extends JpaRepository<KX, Long> {
    @Query(value= "select * from import_fl.list_peasant_economy where iin = ?1", nativeQuery = true)
    List<KX> getKxIin(String IIN);
}
