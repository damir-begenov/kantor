package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.QoldauSubsidy;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface QoldauRepo extends JpaRepository<QoldauSubsidy, String> {
    @Query(value = "SELECT * FROM imp_kfm_fl.qoldau_subsidy where bin = ?1", nativeQuery = true)
    List<QoldauSubsidy> getByIIN(String iin);
}