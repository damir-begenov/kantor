package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.FlRelatives;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface FlRelativesRepository extends JpaRepository<FlRelatives, Long>  {
    @Query(value = "select d.ru_name as name ,*, d.test_column_relation from public.relations_3_level r left join dictionary.d_relations d on d.id = CAST(r.\"RELATION\" AS INT) where \"REF_IIN\" = ?1",nativeQuery = true)
    List<Object[]> findAllByIin(String iin);
    @Query(value= "select * from imp_zags.fl_relations_3_level where parent_iin = ?1", nativeQuery = true)
    List<FlRelatives> getRelativesByFio(String IIN);
}
