package kz.dossier.repositoryDossier;


import kz.dossier.modelsDossier.Credentials;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.Optional;

public interface CredentialsRepo extends JpaRepository<Credentials, String> {
    Optional<Credentials> findByLogin(String login);
//    Boolean existsByUsername(String login);
//
//    Boolean existsByEmail(String login);
}
