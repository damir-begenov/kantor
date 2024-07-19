package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Address;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node
public class REG_ADDRESS_HIST {
    @Id
    @GeneratedValue
    public Long id;
    @Property("Вид связи")
    public String Vid_svyaziey;
    @TargetNode
    private Address address;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getVid_svyaziey() {
        return Vid_svyaziey;
    }

    public void setVid_svyaziey(String vid_svyaziey) {
        Vid_svyaziey = vid_svyaziey;
    }
    public Address getAddress() {
        return address;
    }

    public void setAddress(Address address) {
        this.address = address;
    }
}