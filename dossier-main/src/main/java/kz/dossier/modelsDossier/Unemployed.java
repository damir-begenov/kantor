package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table
@Data
public class Unemployed {
    @Id
    private Long id;
    private String bin;
    private String iin;
    private String profession;
    private String finance_source;
    private String information_source;
    private String completion_status;
    private String salary;
}
