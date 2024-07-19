package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.MillitaryAccount;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;


@Repository
public interface militaryAccountRepo extends JpaRepository<MillitaryAccount, Long> {
    List<MillitaryAccount> findAllByIin(String iin);
}
