package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Company;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node
public class PDL {
    @Id
    @GeneratedValue
    public Long id;
    public String Label;
    public String Source;
    @Property("Вид связи")
    public String Vid_svyaziey;
    @Property("Должность")
    public String Dolzhnost;
    @Property("Орган")
    public String Organ;

    @TargetNode
    private Company company;

    public String getOrgan() {
        return Organ;
    }

    public void setOrgan(String organ) {
        Organ = organ;
    }

    public String getDolzhnost() {
        return Dolzhnost;
    }

    public void setDolzhnost(String dolzhnost) {
        Dolzhnost = dolzhnost;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getLabel() {
        return Label;
    }

    public void setLabel(String label) {
        Label = label;
    }

    public String getSource() {
        return Source;
    }

    public void setSource(String source) {
        Source = source;
    }

    public String getVid_svyaziey() {
        return Vid_svyaziey;
    }

    public void setVid_svyaziey(String vid_svyaziey) {
        Vid_svyaziey = vid_svyaziey;
    }

    public Company getCompany() {
        return company;
    }

    public void setCompany(Company company) {
        this.company = company;
    }
}
