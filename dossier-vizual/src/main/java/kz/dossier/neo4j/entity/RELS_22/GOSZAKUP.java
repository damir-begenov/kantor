package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Company;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node
public class GOSZAKUP {
    @Id
    @GeneratedValue
    public Long id;
    @Property("Вид связи")
    public String Vid_svyaziey;
    @Property("Описание")
    public String Description;
    @TargetNode
    private Company company;

    public String getDescription() {
        return Description;
    }

    public void setDescription(String description) {
        Description = description;
    }

    public Company getCompany() {
        return company;
    }

    public void setCompany(Company company) {
        this.company = company;
    }

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

}
