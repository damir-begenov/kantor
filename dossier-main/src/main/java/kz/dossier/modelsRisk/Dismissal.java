package kz.dossier.modelsRisk;


import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.transaction.annotation.Transactional;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Table(name = "dismissal", schema = "imp_risk")
@Transactional(readOnly = true)
public class Dismissal {
    @Id
    private Integer id;
    private String organ;
    private String punishment;
    private String iin;
    private String work_place;
    private String order_date;
    private String order_num;

}
