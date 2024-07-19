package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvUl;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository

public interface MvUlRepo extends JpaRepository<MvUl, Long> {
    @Query(value= "select * from initial_data.ul mv_ul0_ where mv_ul0_.bin = ?1 limit 1 ", nativeQuery = true)
    List<MvUl> getUsersByLike(String iin);
    @Query(value= "select DISTINCT ON (bin) * from initial_data.ul where UPPER (full_name_rus) like ?1 or UPPER (full_name_kaz) like ?1", nativeQuery = true)
    List<MvUl> getUlsByName(String name);
    @Query(value="select short_name from initial_data.ul where bin = ?1 limit 1", nativeQuery = true)
    String getNameByBin(String bin);
}