package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.BailiffListEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface BailiffListEntityRepo extends JpaRepository<BailiffListEntity, Long> {
    @Query(value = "select * from imp_notary.bailiff_list  where iin = ?1", nativeQuery = true)
    List<BailiffListEntity> getUsersByLike(String iin);
}