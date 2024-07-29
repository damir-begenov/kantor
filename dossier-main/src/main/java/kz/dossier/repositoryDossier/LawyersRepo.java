package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.Lawyers;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository
public interface LawyersRepo extends JpaRepository<Lawyers, Integer> {
    @Query(value = "select * from imp_kfm_fl.lawyers where iin = ?1", nativeQuery = true)
    List<Lawyers> getByIin(String iin);
}
