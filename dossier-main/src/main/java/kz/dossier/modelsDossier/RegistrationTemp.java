package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Getter;
import lombok.Setter;

@Entity
@Table(name = "registration_temp", schema = "import_fl")
@Getter
@Setter
public class RegistrationTemp {
    @Column(name = "iin")
    private String iin;
    @Id
    @Column(name = "person_id")
    private String personId;

    @Column(name = "country")
    private String country;

    @Column(name = "region")
    private String region;

    @Column(name = "district")
    private String district;

    @Column(name = "city")
    private String city;

    @Column(name = "street")
    private String street;

    @Column(name = "building")
    private String building;

    @Column(name = "flat")
    private String flat;

    @Column(name = "address_type_id")
    private String addressTypeId;

    @Column(name = "region_name")
    private String regionName;

    @Column(name = "district_name")
    private String districtName;

    @Column(name = "corpus")
    private String corpus;

    @Column(name = "begin_date")
    private String beginDate;

    @Column(name = "end_date")
    private String endDate;

    @Column(name = "ar_code")
    private String arCode;

    @Column(name = "id")
    private String id;
}
