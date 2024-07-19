package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Persons;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node("BERKUT")
public class BERKUT {
    @Id
    @GeneratedValue
    public Long id;

    @Property("Описание")
    public String Description;

    @Property("Вид связи")
    public String Vid_svyaziey;

    @TargetNode
    private Persons person;

    public Persons getPerson() {
        return person;
    }

    public void setPerson(Persons person) {
        this.person = person;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getDescription() {
        return Description;
    }

    public void setDescription(String description) {
        Description = description;
    }

    public String getVid_svyaziey() {
        return Vid_svyaziey;
    }

    public void setVid_svyaziey(String vid_svyaziey) {
        Vid_svyaziey = vid_svyaziey;
    }
}
