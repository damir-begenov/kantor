package kz.dossier.dto;

import kz.dossier.modelsDossier.FpgTempEntity;
import kz.dossier.modelsDossier.Msh;
import kz.dossier.modelsRisk.NdsEntity;
import kz.dossier.modelsRisk.*;
import kz.dossier.modelsRisk.OpgEntity;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.List;
@Data
public class UlRiskDTO {
    private String bin;
    private List<Bankrot> bankrots; //--Принудительное лечение
    private List<BlockEsf> blockEsfs;
    private List<Criminals> criminals; //--Банкроты
    private List<Dormant> dormants; //--Участие в рисковых ЮЛ
    private List<Kuis> kuis; //--Участие в рисковых ЮЛ
    private List<FirstCreditBureauEntity> firstCreditBureauEntities; //--Сироты
    private List<NdsEntity> ndsEntities; //--1Д
    private List<Omn> omns;
    private List<OpgEntity> opgEntities; //-- Cписок льготников
    private List<FpgTempEntity> fpgTempEntities;
    private List<VerificationField> verificationFields;
    private Double percentage;
    private int quantity; //--количество

    public void setQuantity(int quantity) {
        int riskListCount = countRiskLists();
        this.percentage = (double) (quantity * 100 / riskListCount);
        this.quantity = quantity;
    }
    private int countRiskLists() {
        Field[] fields = FLRiskDto.class.getDeclaredFields();
        int count = 0;
        for (Field field : fields) {
            if (List.class.isAssignableFrom(field.getType())) {
                count++;
            }
        }
        return count;
    }
}
