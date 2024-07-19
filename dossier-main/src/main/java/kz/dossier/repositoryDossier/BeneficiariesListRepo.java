package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.BeneficiariesList;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository
public interface BeneficiariesListRepo extends JpaRepository<BeneficiariesList, Integer> {
    @Query(value = "Select * from imp_risk.beneficiaries_list where iin = ?1", nativeQuery = true)
    List<BeneficiariesList> getBenByIIN(String IIN);

}
