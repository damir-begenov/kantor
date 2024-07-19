package kz.dossier.repositoryDossier;


import kz.dossier.modelsRisk.Bankrot;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface BankrotRepo extends JpaRepository<Bankrot, Long> {
    @Query(value= "select * from imp_risk.bankrot where iin_bin = ?1", nativeQuery = true)
    List<Bankrot> getbankrotByByIIN(String IIN);
}
