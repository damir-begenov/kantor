package kz.dossier.modelsDossier;


import com.fasterxml.jackson.annotation.JsonPropertyOrder;
import kz.dossier.modelsRisk.*;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.List;

@Data
@JsonPropertyOrder({ "flCard","criminalRecordDtos","convictsRiskList","minsdravInfoList","compulsaryTreatmentRiskDtos",
        "orphansRawDtoList","omnDtoList","dormantsDtoList","removedAccountNdsRiskDtos","blockingEcfList","kgdBankruptsList",
        "pfrList","drugAddictsRisks","erdrUdCardDtoList","ulParticipantInfo","iin","quantity","percentage"})
public class FLRiskDto {

    private String iin;
    private List<AccountantListEntity> accountantListEntities; //--Осужденные
    private List<Adm> adms; //--Сведения Минздрава
    private List<Bankrot> bankrots; //--Принудительное лечение
    private List<BlockEsf> blockEsfs;
    private List<ConvictsJustified> convictsJustifieds; //--Снятые с учета по НДС
    private List<ConvictsTerminatedByRehab> convictsTerminatedByRehabs; //--Блокировка выписки ЭСФ
    private List<ConvictsAbroad> convictsAbroads; //--Блокировка выписки ЭСФ
    private List<Criminals> criminals; //--Банкроты
    private List<Dismissal> dismissals; //--Банкроты
    private List<Dormant> dormants; //--Участие в рисковых ЮЛ
    private List<Kuis> kuis; //--Участие в рисковых ЮЛ
    private List<Incapacitated> incapacitateds; //--Участие в рисковых ЮЛ
    private List<FirstCreditBureauEntity> firstCreditBureauEntities; //--Сироты
    private List<FpgTempEntity> fpgTempEntities; //--Бездействующие
    private List<ImmoralLifestyle> immoralLifestyles; //--Отсутствие по месту нахождения(ОМН)
    private List<Msh> mshes;//--карточки
    private List<MzEntity> mzEntities; //--Осужденныен по статьям
    private List<NdsEntity> ndsEntities; //--1Д
    private List<Omn> omns;
    private List<OpgEntity> opgEntities; //-- Cписок льготников
    private List<Orphans> orphans; //-- Увольнения по отрицательным мотивам
    private List<BeneficiariesList> beneficiariesLists;
    private List<WantedListEntity> wantedListEntities;
    private List<DrugAddicts> drugAddicts;
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
