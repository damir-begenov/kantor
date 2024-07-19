package kz.dossier.modelsRisk;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
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
