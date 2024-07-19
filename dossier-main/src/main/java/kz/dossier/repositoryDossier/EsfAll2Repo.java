package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.EsfAll2;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface EsfAll2Repo extends JpaRepository<EsfAll2, Long> {
//    @Query(value = "select * from esf_all2 limit 1", nativeQuery = true)
//    List<esf_all2> findRoleById(Long id);
}
