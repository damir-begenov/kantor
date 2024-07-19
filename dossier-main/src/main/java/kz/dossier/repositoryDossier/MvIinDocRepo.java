package kz.dossier.repositoryDossier;

import kz.dossier.modelsDossier.MvIinDoc;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.UUID;
@Repository
public interface MvIinDocRepo extends JpaRepository<MvIinDoc, UUID> {
    @Query(value = "select * from initial_data.fl_iin_doc where (iin = ?1 and expiry_date = (select max(expiry_date) from initial_data.fl_iin_doc where " +
            "iin = ?1 and doc_type_ru_name = 'УДОСТОВЕРЕНИЕ РК')) or (iin = ?1 and doc_type_ru_name != 'УДОСТОВЕРЕНИЕ РК' and expiry_date=(select max(expiry_date)" +
            "from initial_data.fl_iin_doc where iin = ?1 and doc_type_ru_name != 'УДОСТОВЕРЕНИЕ РК'))", nativeQuery = true)
    List<MvIinDoc> getByIIN(String iin);

    @Query(value = "SELECT iin FROM initial_data.fl_iin_doc WHERE doc_number = ?1 LIMIT 1", nativeQuery = true)
    String getIinByDoc_Number(String doc);
    @Query(value = "select * FROM initial_data.fl_iin_doc where doc_number = ?1 ", nativeQuery = true)
    List<MvIinDoc> getByDoc_number(String doc_number);
}
