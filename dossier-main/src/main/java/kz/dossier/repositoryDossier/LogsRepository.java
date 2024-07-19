package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.Logs;
import org.springframework.data.jpa.repository.JpaRepository;

import java.time.LocalDateTime;

public interface LogsRepository extends JpaRepository<Logs, LocalDateTime> {

}
