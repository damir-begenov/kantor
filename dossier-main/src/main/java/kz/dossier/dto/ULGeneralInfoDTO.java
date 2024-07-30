package kz.dossier.dto;

import kz.dossier.modelsDossier.CommodityProducer;
import kz.dossier.modelsDossier.FlContacts;
import kz.dossier.modelsDossier.MvUl;
import kz.dossier.modelsDossier.RegAddressUlEntity;
import kz.dossier.modelsDossier.AccountantListEntity;
import kz.dossier.modelsRisk.Pdl;
import lombok.Data;

import java.util.List;
import java.util.Optional;

@Data
public class ULGeneralInfoDTO {
    private Double percentage;
    private Optional<MvUl> mvUlList;
    private RegAddressUlEntity regAddressUlEntity;
    private List<Pdl> pdls;
    private List<CommodityProducer> commodityProducers;
    private List<AccountantListEntity> accountantListEntities;
    private List<FlContacts> flContacts;
}
