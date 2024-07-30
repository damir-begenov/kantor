package kz.dossier.modelsRisk;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Data
@Entity
@Table(name = "sudispol", schema = "imp_risk")
public class Sudispol {
    @Id
    private Long id;
    private String name_of_organ;
    private String name_of_layer;
    private String number_of_man;
    private String date_initiation;
    private String category_requirements;
    private String addition_category;
    private String amount;
}
