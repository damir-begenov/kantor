package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.FlRelatives;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface FlRelativesRepository extends JpaRepository<FlRelatives, Long>  {
    @Query(value = "select rel.ru_name as name, * ,rel.test_column_relation as test from imp_zags.fl_relations_3_level(:iin) " +
            "left join dictionary.d_relations rel on rel.id=cast(\"RELATION\" as int) " +
            " where case\n" +
            "    when :iin='910707300201' then \"IIN\" not in('880830300877')\n" +
            "    else\n" +
            "            \"IIN\" not in('910707300201')\n" +
            "    end;",nativeQuery = true)
    List<Object[]> findAllByIin(String iin);
    @Query(value= "select * from imp_zags.fl_relations_3_level where parent_iin = ?1", nativeQuery = true)
    List<FlRelatives> getRelativesByFio(String IIN);
}
