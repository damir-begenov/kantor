package kz.dossier.modelsDossier;

import lombok.Data;

import java.util.List;

@Data
public class FlFirstRowDto {
    private List<MvFl> mvFls;
    private List<PhotoDb> photoDbf;
    private List<MvIinDoc> mvIinDocList;
    private List<MvFlAddress> mvFlAddresses;
    private List<RegistrationTemp> registrationTemps;
    private double riskPercentage;
}
