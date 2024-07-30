package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvUlFounderUl;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository

public interface MvUlFounderUlRepo extends JpaRepository<MvUlFounderUl, Long> {
    @Query(value= "select * from imp_kfm_ul.ul_founder_ul  where bin_org = ?1  ", nativeQuery = true)
    List<MvUlFounderUl> getUsersByLike(String iin);
}