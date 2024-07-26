package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvFlAddress;
import kz.dossier.modelsDossier.RegistrationTemp;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface RegistrationTempRepository extends JpaRepository<RegistrationTemp, Long> {
    @Query(value = "select * from import_fl.registration_temp where iin = ?1", nativeQuery = true)
    List<RegistrationTemp> getRegAddressByIIN(String iin);

}
