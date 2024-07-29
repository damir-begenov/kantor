package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table(name = "auto_transport", schema = "imp_kfm_fl")
@Data
public class AutoTransport {
    @Id
    private Long id;
    private String iin_bin;
    private String brand;
    private String number;
    private String date;

}
