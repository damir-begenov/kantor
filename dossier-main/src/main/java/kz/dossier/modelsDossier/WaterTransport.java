package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table(name = "water_transport", schema = "imp_kfm_fl")
@Data
public class WaterTransport {
    private String ikt;
    private String reg_number;
    private String iin_bin;
    private String name;
    private String year_reestr;
    private String name_vt;
    private String type_vt;
    private String purpose;
    private String region_vt;
    private String material;
    private String country_vt;
    private String name_engine;
    private String sourse;
    @Id
    private Long id;
}

