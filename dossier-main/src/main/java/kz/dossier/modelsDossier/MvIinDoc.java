package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

import java.sql.Date;
import java.util.UUID;

@Entity
@Table(name = "mv_fl_documents", schema = "imp_kfm_fl")
@Data
public class MvIinDoc {
    @Id
    private UUID id;
    private String iin;
    private String doc_number;
    private String doc_type_id;
    private String doc_type_ru_name;
    private Date issue_date;
    private Date expiry_date;
    private String document_invalidity_id;
    private String document_invalidity_ru_name;
    private Date document_invalidity_date;
    private String issue_organization_id;
    private String issue_organization_ru_name;

}
