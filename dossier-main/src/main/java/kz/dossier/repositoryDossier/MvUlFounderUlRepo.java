package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvUlFounderUl;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository

public interface MvUlFounderUlRepo extends JpaRepository<MvUlFounderUl, Long> {
    @Query(value= "select * from initial_data.ul_founder_ul mv_ul0_ where mv_ul0_.bin_org = ?1  ", nativeQuery = true)
    List<MvUlFounderUl> getUsersByLike(String iin);
}