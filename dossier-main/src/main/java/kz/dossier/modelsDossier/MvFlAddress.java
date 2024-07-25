package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Getter;
import lombok.Setter;

import java.util.UUID;

@Entity
@Table(name = "mv_fl_address", schema = "imp_kfm_fl")
@Getter
@Setter
public class MvFlAddress {
    @Id
    @Column(name = "ID")
    private UUID id;

    @Column(name = "IIN")
    private String iin;

    @Column(name = "fio")
    private String fio;

    @Column(name = "ADDRESS_DATE")
    private String addressDate;

    @Column(name = "ADDRESS_END_DATE")
    private String addressEndDate;

    @Column(name = "RKA")
    private String rka;

    @Column(name = "RN_ADDRESS_RU")
    private String rnAddressRu;

    @Column(name = "RN_ADDRESS_HISTORY_RU")
    private String rnAddressHistoryRu;

    @Column(name = "DISTRICT")
    private String district;

    @Column(name = "CITY")
    private String city;

    @Column(name = "REGION")
    private String region;

    @Column(name = "STREET")
    private String street;

    @Column(name = "BUILDING")
    private String building;

    @Column(name = "CORPUS")
    private String corpus;

    @Column(name = "FLAT")
    private String flat;

}
