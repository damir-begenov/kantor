package kz.dossier.modelsRisk;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

import java.util.Date;

@Entity
@Table
@Data
public class VerificationField {
    @Id
    private Long id;
    private String verification_organ;
    private String verification_result;
    private String scope_of_control;
    private String basis_of_verification;
    private String verification_type;
    private String bin;
    private Date reg_date;
}
