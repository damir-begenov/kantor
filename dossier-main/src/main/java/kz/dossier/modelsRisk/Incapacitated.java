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
@Table(name = "incapacitated", schema = "imp_risk")
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Incapacitated {
    @Id
    private Integer id;
    private String iin;
    private String court;
    private Date date_decision;
    private String soct;
    private String cause;
}
