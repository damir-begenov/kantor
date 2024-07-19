package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "imp_kfm_fl.ipgo_email")
public class IpgoEmailEntity {
    @Id    @Nullable

    private Long id;

    @Nullable
    @Column(name = "iin")
    private String iin;

    @Nullable
    @Column(name = "organ")
    private String organ;

    @Nullable
    @Column(name = "position")
    private String position;

    @Column(name = "email")
    @Nullable
    private String email;

    @Nullable
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
    public String getOrgan() {
        return organ;
    }

    public void setOrgan(@Nullable String organ) {
        this.organ = organ;
    }

    @Nullable
    public String getPosition() {
        return position;
    }

    public void setPosition(@Nullable String position) {
        this.position = position;
    }

    @Nullable
    public String getEmail() {
        return email;
    }

    public void setEmail(@Nullable String email) {
        this.email = email;
    }
}