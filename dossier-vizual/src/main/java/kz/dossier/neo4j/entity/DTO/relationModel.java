package kz.dossier.neo4j.entity.DTO;

import java.util.Map;

public class relationModel {
    private Long from;
    private Long to;
    private String type;
    private Map<String, Object> properties;

    public relationModel(Long start, Long end, Map<String, Object> propertiesModels) {
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

    public Long getFrom() {
        return from;
    }

    public void setFrom(Long from) {
        this.from = from;
    }

    public Long getTo() {
        return to;
    }

    public void setTo(Long to) {
        this.to = to;
    }

    public void setType(String type) {
        this.type = type;
    }
    public void setProperties(Map<String, Object> properties) {
        this.properties = properties;
    }
}
