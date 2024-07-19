package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.RegAddressFl;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
@Repository

public interface RegAddressFlRepo extends JpaRepository<RegAddressFl, String> {
    @Query(value = "SELECT * FROM initial_data.fl_reg_address where iin = ?1", nativeQuery = true)
    List<RegAddressFl> getByIIN(String iin);
}