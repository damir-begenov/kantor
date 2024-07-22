package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "imp_risk.commodity_producers")
public class CommodityProducer {
    @Id
    @Nullable
    private Integer id;

    @Nullable
    @Column(name = "ssp_name")
    private String sspName;
    @Nullable

    @Column(name = "cnt")
    private Integer count;
    @Nullable

    @Column(name = "iin_bin")
    private String iinBin;
    @Nullable

    @Column(name = "producer")
    private String producer;
    @Nullable

    @Column(name = "status")
    private String status;
    @Nullable

    @Column(name = "region")
    private String region;
    @Nullable

    @Column(name = "sztp")
    private String sztp;

    @Nullable
    public Integer getId() {
        return id;
    }

    public void setId(@Nullable Integer id) {
        this.id = id;
    }

    @Nullable
    public String getSspName() {
        return sspName;
    }

    public void setSspName(@Nullable String sspName) {
        this.sspName = sspName;
    }

    @Nullable
    public Integer getCount() {
        return count;
    }

    public void setCount(@Nullable Integer count) {
        this.count = count;
    }

    @Nullable
    public String getIinBin() {
        return iinBin;
    }

    public void setIinBin(@Nullable String iinBin) {
        this.iinBin = iinBin;
    }

    @Nullable
    public String getProducer() {
        return producer;
    }

    public void setProducer(@Nullable String producer) {
        this.producer = producer;
    }

    @Nullable
    public String getStatus() {
        return status;
    }

    public void setStatus(@Nullable String status) {
        this.status = status;
    }

    @Nullable
    public String getRegion() {
        return region;
    }

    public void setRegion(@Nullable String region) {
        this.region = region;
    }

    @Nullable
    public String getSztp() {
        return sztp;
    }

    public void setSztp(@Nullable String sztp) {
        this.sztp = sztp;
    }
}