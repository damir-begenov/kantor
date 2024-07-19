package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.Universities;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface UniversitiesRepo extends JpaRepository<Universities, Integer> {
    @Query(value = "select * from public.study where iin = ?1", nativeQuery = true)
    List<Universities> getByIIN(String iin);
}
