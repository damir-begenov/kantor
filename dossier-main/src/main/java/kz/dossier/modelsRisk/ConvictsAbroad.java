package kz.dossier.modelsRisk;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "convicts_abroad", schema = "imp_risk")
@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class ConvictsAbroad {
    @Id
    private Integer id;
    private String measure;
    private String iin;
    private Date entry_date;
    private Date court_decision_date;
    private String sud_organ;
    private String qualification;
    private String country;
    private String number;
}
