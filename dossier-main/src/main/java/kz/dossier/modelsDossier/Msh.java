package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "msh", schema = "imp_kfm_fl")
public class Msh {



        @Column(name = "owner_iin_bin")
        @Nullable
        private String ownerIinBin;
        @Nullable

        @Column(name = "equipment_type")
        private String equipmentType;
        @Nullable

        @Column(name = "equipment_model")
        private String equipmentModel;
       @Nullable
       @Id

        @Column(name = "vin")
        private String vin;

    @Nullable
        @Column(name = "gov_number")
        private String govNumber;

    @Nullable
        @Column(name = "reg_series_num")
        private String regSeriesNum;

        @Column(name = "reg_date")
        @Nullable
        private Date regDate;


    @Nullable
    public String getOwnerIinBin() {
        return ownerIinBin;
    }

    public void setOwnerIinBin(@Nullable String ownerIinBin) {
        this.ownerIinBin = ownerIinBin;
    }

    @Nullable
    public String getEquipmentType() {
        return equipmentType;
    }

    public void setEquipmentType(@Nullable String equipmentType) {
        this.equipmentType = equipmentType;
    }

    @Nullable
    public String getEquipmentModel() {
        return equipmentModel;
    }

    public void setEquipmentModel(@Nullable String equipmentModel) {
        this.equipmentModel = equipmentModel;
    }

    @Nullable
    public String getVin() {
        return vin;
    }

    public void setVin(@Nullable String vin) {
        this.vin = vin;
    }

    @Nullable
    public String getGovNumber() {
        return govNumber;
    }

    public void setGovNumber(@Nullable String govNumber) {
        this.govNumber = govNumber;
    }

    @Nullable
    public String getRegSeriesNum() {
        return regSeriesNum;
    }

    public void setRegSeriesNum(@Nullable String regSeriesNum) {
        this.regSeriesNum = regSeriesNum;
    }

    @Nullable
    public Date getRegDate() {
        return regDate;
    }

    public void setRegDate(@Nullable Date regDate) {
        this.regDate = regDate;
    }
}
