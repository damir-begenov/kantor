package kz.dossier.modelsDossier;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Table
@Entity
@Data
public class IndividualEntrepreneur {
    @Id
    private String rnn;
    private String iin;
    private String name_rus;
}
