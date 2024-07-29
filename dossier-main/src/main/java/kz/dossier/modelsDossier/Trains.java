package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

import java.util.Date;

@Entity
@Table
@Data
public class Trains {
    @Id
    private Long id;
    private String iin_bin;
    private Date doc_date;
    private String vagon_type;
    private String vagon_series;
    private String vagon_category;
    private String vagon_make_year;
    private String ownership_type;
}
