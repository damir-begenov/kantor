package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvUlLeader;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;
import java.util.UUID;

public interface MvUlLeaderRepository extends JpaRepository<MvUlLeader, UUID> {
    List<MvUlLeader> findAllByIin(String iin);

    @Query(name = "select * from imp_kfm_ul.ul_leader where bin_org = ?1", nativeQuery = true)
    List<MvUlLeader> findAllByBinOrg(String bin);
}
