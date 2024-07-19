package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Persons;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node(value = "WORKER_HIST")
public class WORKER_HIST {
    @Id
    @GeneratedValue
    public Long id;
    @Property("Вид связи")
    public String Vid_svyaziey;
    @Property("Дата начало работы")
    public String Start_Date;
    @Property("Дата окончания работы")
    public String End_Date;
    @Property("Дата первой зарплаты")
    public String first_payment_date;
    @Property("Дата последней зарплаты")
    public String last_payment_date;

    public String getFirst_payment_date() {
        return first_payment_date;
    }

    public void setFirst_payment_date(String first_payment_date) {
        this.first_payment_date = first_payment_date;
    }

    public String getLast_payment_date() {
        return last_payment_date;
    }

    public void setLast_payment_date(String last_payment_date) {
        this.last_payment_date = last_payment_date;
    }

    @TargetNode
    private Persons person;

    public String getStart_Date() {
        return Start_Date;
    }

    public void setStart_Date(String start_Date) {
        Start_Date = start_Date;
    }

    public String getEnd_Date() {
        return End_Date;
    }

    public void setEnd_Date(String end_Date) {
        End_Date = end_Date;
    }

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

    public String getVid_svyaziey() {
        return Vid_svyaziey;
    }

    public void setVid_svyaziey(String vid_svyaziey) {
        Vid_svyaziey = vid_svyaziey;
    }
}
