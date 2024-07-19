package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.FlPensionMini;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface FlPensionMiniRepo extends JpaRepository<FlPensionMini, Long> {


    @Query(value = "SELECT EXTRACT(YEAR FROM \"PAY_DATE\") AS PAY_DATE,\n" +
            "       \"P_NAME\", \"KNP\",\n" +
            "    SUM(\"AMOUNT\") AS \"AMOUNT\"\n" +
            "    FROM imp_kfm_fl.fl_pension_contr \n" +
            "    WHERE \"IIN\" = ?1 and \"P_RNN\" = ?2 \n" +
            "    GROUP BY EXTRACT(YEAR FROM \"PAY_DATE\"), \"P_NAME\", \"KNP\"  ", nativeQuery = true)
    List<FlPensionMini> getAllByCompanies(String iin , String bin);
    @Query(value = "SELECT DISTINCT(EXTRACT(YEAR FROM \"PAY_DATE\")) \n" +
            "    FROM imp_kfm_fl.fl_pension_contr \n" +
            "    WHERE \"IIN\" = ?1 and \"P_RNN\" = ?2 \n", nativeQuery = true)
    List<String> getAllByCompaniesYear(String iin , String bin);
    @Query(value = "select distinct extract(year from \"PAY_DATE\")  from imp_kfm_fl.fl_pension_contr fpc where \"IIN\" = ?1", nativeQuery = true)
    List<String> getYears(String iin);
}
