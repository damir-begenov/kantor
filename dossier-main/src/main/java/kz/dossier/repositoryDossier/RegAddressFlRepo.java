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

    @Query(value = "SELECT * FROM initial_data.fl_reg_address where iin = ?1 and registration_type = 'Permanent' order by reg_date desc limit 1", nativeQuery = true)
    List<RegAddressFl> getByPermanentIin(String iin);

    @Query(value = "SELECT * FROM initial_data.fl_reg_address where region = ?1 and district = ?2 and city = ?3 and street = ?4 and building = ?5 and korpus = ?6 and apartment_number = ?7", nativeQuery = true)
    List<RegAddressFl> getByAddress(String region, String district, String city, String street, String building, String korpus, String apartment_number);
}