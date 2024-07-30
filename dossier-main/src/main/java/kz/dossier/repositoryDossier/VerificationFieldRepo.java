package kz.dossier.repositoryDossier;

import kz.dossier.modelsRisk.WantedListEntity;
import kz.dossier.modelsRisk.VerificationField;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface VerificationFieldRepo extends JpaRepository<VerificationField,Long> {
    @Query(value = "SELECT * FROM imp_risk.verification_filed where bin = ?1", nativeQuery = true)
    List<VerificationField> getByIIN(String iin);
}
