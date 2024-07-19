package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvUlLeader;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.UUID;

public interface MvUlLeaderRepository extends JpaRepository<MvUlLeader, UUID> {
    List<MvUlLeader> findAllByIin(String iin);
}
