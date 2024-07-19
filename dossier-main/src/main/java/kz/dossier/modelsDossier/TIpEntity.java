package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "imp_kfm_fl.mv_ip")
public class TIpEntity {
    @Id
    @Nullable

    private Long id;
    @Nullable

    @Column(name = "iin")
    private String iin;
    @Nullable

    @Column(name = "name")
    private String ip_name;

    @Nullable
    @Column(name = "fio")
    private String fio;

    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getName() {
        return ip_name;
    }

    public void setName(@Nullable String name) {
        this.ip_name = name;
    }

    @Nullable
    public String getFio() {
        return fio;
    }

    public void setFio(@Nullable String fio) {
        this.fio = fio;
    }
}