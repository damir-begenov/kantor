package kz.dossier.neo4j.entity.RELS_22;

import kz.dossier.neo4j.entity.Company;
import org.springframework.data.neo4j.core.schema.*;

@RelationshipProperties
@Node
public class operation {
    @Id
    @GeneratedValue
    public Long id;
    @Property("amount")
    public String amount;
    @Property("amount_tenge")
    public String amount_tenge;
    @Property("cfm_name")
    public String cfm_name;
    @Property("date")
    public String date;
    @Property("fm_message")
    public String fm_message;
    @Property("oper_type")
    public String oper_type;

    @TargetNode
    private Company company;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getAmount() {
        return amount;
    }

    public void setAmount(String amount) {
        this.amount = amount;
    }

    public String getAmount_tenge() {
        return amount_tenge;
    }

    public void setAmount_tenge(String amount_tenge) {
        this.amount_tenge = amount_tenge;
    }

    public String getCfm_name() {
        return cfm_name;
    }

    public void setCfm_name(String cfm_name) {
        this.cfm_name = cfm_name;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getFm_message() {
        return fm_message;
    }

    public void setFm_message(String fm_message) {
        this.fm_message = fm_message;
    }

    public String getOper_type() {
        return oper_type;
    }

    public void setOper_type(String oper_type) {
        this.oper_type = oper_type;
    }

    public Company getCompany() {
        return company;
    }

    public void setCompany(Company company) {
        this.company = company;
    }
}
