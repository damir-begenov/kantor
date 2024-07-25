package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Id;
import lombok.Data;
import org.jetbrains.annotations.Nullable;

@Data
public class MilitaryAccountingDTO {
    private Long id;
    private String bin;
    private String dateStart;
    private String iin;
    private String dateEnd;
    private String binName;
}

