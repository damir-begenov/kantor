package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.IndividualEntrepreneur;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface IndividualEntrepreneurRepo extends JpaRepository<IndividualEntrepreneur, Long> {
    @Query(value = "select * from import_fl.list_individual_entrepreneur where iin = ?1", nativeQuery = true)
    List<IndividualEntrepreneur> getByIin(String iin);
}
