package kz.dossier.modelsRisk;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "drug_addicts_pfr", schema = "imp_risk")
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class DrugAddicts {
    @Id
    private Integer id;
    private String source;
    private String fio;
    private String iin;
    private String ishod;
}
