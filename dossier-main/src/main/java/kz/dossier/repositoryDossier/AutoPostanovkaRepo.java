package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.AutoPostanovka;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface AutoPostanovkaRepo extends JpaRepository<AutoPostanovka, Long> {
    @Query(value = "select * from transport.постановка where \"ИИН\" = ?1 ", nativeQuery = true)
    List<AutoPostanovka> getAutoPostanovkaByIin(String iin);
}
