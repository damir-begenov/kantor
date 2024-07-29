package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

import java.util.Date;

@Entity
@Table
@Data
public class AviaTransport {
    @Id
    private Long id;
    private Date date_exclusion;
    private String ekspluatant;
    private String iin;
    private String registraion_number;
    private String bort;
    private String marka;
    private Date date_registration;
}
