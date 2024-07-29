package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table(name ="list_peasant_economy", schema = "import_fl")
@Data
public class KX {
    @Id
    private String rnn;
    private String iin;
    private String name_rus;
}
