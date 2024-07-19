package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.AdvocateListEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface AdvocateListEntityRepo extends JpaRepository<AdvocateListEntity, Long> {
    @Query(value = "select * from imp_notary.advocate_list  where iin = ?1", nativeQuery = true)
    List<AdvocateListEntity> getUsersByLike(String iin);
}