package kz.dossier.modelsDossier;

import lombok.Data;

import java.sql.Blob;
import java.util.List;

@Data
public class FlRelativesLevelDto {
    //--Сведения о родственных связях данного ФЛ - Визуализация данных
    private String name;
    private String iin;
    private String fio;
    private Blob photo;
    //    private String relation;
//    private String birthDate;
//    private String marriageRegDate;
//    private String marriageDivorceDate;
    private Boolean haveRisk;
    private Boolean isDirector;
    private Boolean isContinue;
    private List<FlRelativesLevelDto> children;
}
