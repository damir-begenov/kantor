package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Address;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node
public class REG_ADDRESS_CUR {
    @Id
    @GeneratedValue
    public String id;
    @Property("Вид связи")
    public String Vid_svyaziey;
    @TargetNode
    private Address address;

    public Address getAddress() {
        return address;
    }

    public void setAddress(Address address) {
        this.address = address;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getVid_svyaziey() {
        return Vid_svyaziey;
    }

    public void setVid_svyaziey(String vid_svyaziey) {
        Vid_svyaziey = vid_svyaziey;
    }
}
