package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.ChangeFio;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.Optional;
@Repository
public interface ChangeFioRepo extends JpaRepository<ChangeFio, String> {
    @Query(value = "select * from imp_zags.change_fio where iin = ?1 limit 1", nativeQuery = true)
    Optional<ChangeFio> getByIin(String iin);
}
