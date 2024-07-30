package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.AutoPostanovka;
import kz.dossier.modelsDossier.AutoSnyatie;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface AutoSnyatieRepo  extends JpaRepository<AutoSnyatie, Long> {
    @Query(value = "select * from transport.снятые where \"ИИН\" = ?1 ", nativeQuery = true)
    List<AutoSnyatie> getAutoSnyatieByIin(String iin);
    @Query(value = "select * from transport.снятые where \"БИН субъекта\" = ?1 ", nativeQuery = true)
    List<AutoSnyatie> getAutoSnyatieByBin(String iin);
}
