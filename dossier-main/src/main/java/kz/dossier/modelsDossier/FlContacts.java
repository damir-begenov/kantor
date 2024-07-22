package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "contacts", schema = "imp_kfm_fl")
public class FlContacts {
    @Nullable
    private String iin;
    @Nullable
    private String phone;
    @Nullable
    private String fio;
    @Nullable
    private String email;
    @Nullable
    private String nickname;
    @Nullable
    private String source;
    @Nullable
    private String leader_fio;
    @Id
    @Nullable
    private long id;

    @Nullable
    public String getIin() {
        return iin;
    }

    public void setIin(@Nullable String iin) {
        this.iin = iin;
    }

    @Nullable
    public String getPhone() {
        return phone;
    }

    public void setPhone(@Nullable String phone) {
        this.phone = phone;
    }

    @Nullable
    public String getFio() {
        return fio;
    }

    public void setFio(@Nullable String fio) {
        this.fio = fio;
    }

    @Nullable
    public String getEmail() {
        return email;
    }

    public void setEmail(@Nullable String email) {
        this.email = email;
    }

    @Nullable
    public String getNickname() {
        return nickname;
    }

    public void setNickname(@Nullable String nickname) {
        this.nickname = nickname;
    }

    @Nullable
    public String getSource() {
        return source;
    }

    public void setSource(@Nullable String source) {
        this.source = source;
    }

    @Nullable
    public String getLeader_fio() {
        return leader_fio;
    }

    public void setLeader_fio(@Nullable String leader_fio) {
        this.leader_fio = leader_fio;
    }

    public long getId() {
        return id;
    }

    public void setId(long id) {
        this.id = id;
    }
}
