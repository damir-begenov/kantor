package kz.dossier.modelsRisk;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Table(name = "beneficiaries_list", schema = "imp_risk")
public class BeneficiariesList {
    @Id
    private Integer id;
    private String iin;
    private String status;
}
