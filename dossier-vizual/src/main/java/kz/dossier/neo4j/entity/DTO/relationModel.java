package kz.dossier.neo4j.entity.DTO;

import java.util.Map;

public class relationModel {
    private String from;
    private String to;
    private String type;
    private Map<String, Object> properties;

    public relationModel(String start, String end, Map<String, Object> propertiesModels) {
        this.from = start;
        this.to = end;
        this.properties = propertiesModels;
    }


    public String getType() {
        return type;
    }
    public Map<String, Object> getProperties() {
        return properties;
    }

    public String getFrom() {
        return from;
    }

    public void setFrom(String from) {
        this.from = from;
    }

    public String getTo() {
        return to;
    }

    public void setTo(String to) {
        this.to = to;
    }

    public void setType(String type) {
        this.type = type;
    }
    public void setProperties(Map<String, Object> properties) {
        this.properties = properties;
    }
}
