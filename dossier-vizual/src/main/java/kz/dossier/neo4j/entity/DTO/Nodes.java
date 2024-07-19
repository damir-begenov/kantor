package kz.dossier.neo4j.entity.DTO;


import kz.dossier.modelsDossier.PhotoDb;

import java.util.Map;

public class  Nodes {
    private Long id;
    private PhotoDb photoDbf;
    private Map<String, Object> properties;

    private Long relCount;

    public Long getRelCount() {
        return relCount;
    }

    public void setRelCount(Long relCount) {
        this.relCount = relCount;
    }

    public PhotoDb getPhotoDbf() {
        return photoDbf;
    }

    public void setPhotoDbf(PhotoDb photoDbf) {
        this.photoDbf = photoDbf;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Map<String, Object> getProperties() {
        return properties;
    }

    public void setProperties(Map<String, Object> properties) {
        this.properties = properties;
    }
}
