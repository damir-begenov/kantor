package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.School;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface SchoolRepo extends JpaRepository<School, Integer> {
    @Query(value = "select * from public.school where iin = ?1", nativeQuery = true)
    List<School> getByIIN(String iin);
}
