package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.RegAddressUlEntity;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

public interface RegAddressUlEntityRepo extends JpaRepository<RegAddressUlEntity, Long> {
    @Query(value= "select * from imp_kfm_ul_old.mv_reg_address_ul mv_ul0_ where mv_ul0_.bin = ?1 ORDER BY reg_date desc limit 1", nativeQuery = true)
    RegAddressUlEntity findByBin(String bin);
    @Query(value= "SELECT  * FROM imp_kfm_ul_old.mv_reg_address_ul where reg_addr_region_ru = ?1 and reg_addr_district_ru = ?2 and reg_addr_locality_ru= ?3 and reg_addr_street_ru = ?4 and reg_addr_bulding_num = ?5 and bin != ?6 order by reg_date desc limit 1", nativeQuery = true)

    RegAddressUlEntity regAddressNaOdnomMeste(String region,String disctrict, String locality, String street, String bulding, String bin);

    @Query(value = "SELECT * FROM initial_data.fl_reg_address WHERE reg_addr_region_ru = ?1 AND reg_addr_district_ru = ?2 AND reg_addr_rural_district_ru = ?3 AND reg_addr_locality_ru = ?4 AND reg_addr_street_ru = ?5 AND reg_addr_bulding_num = ?6 AND reg_addr_block_num = ?7 AND reg_addr_builing_body_num = ?8 AND reg_addr_office = ?9", nativeQuery = true)
    List<RegAddressUlEntity> getByAddress(String reg_addr_region_ru, String reg_addr_district_ru, String reg_addr_rural_district_ru, String reg_addr_locality_ru, String reg_addr_street_ru, String reg_addr_bulding_num, String reg_addr_block_num, String reg_addr_builing_body_num, String reg_addr_office);

}