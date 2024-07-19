package kz.dossier.modelsRisk;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "kuis", schema = "imp_risk")
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Kuis {
    @Id
    private Integer id;
    private String iin;
    private String institution;
    private String region;
    private String bin;
    private String arrival_date;
    private String period_start;
    private String period_end;
    private String personal_closure_date;
    private String personal_closure_ground;
    private String departure_date;
    private String conviction_article;
    private String editorial_code;

}
