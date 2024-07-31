package kz.dossier.dto;

import java.util.List;

public class PensionGroupDTO {
    private String name;
    private List<PensionListDTO> list;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<PensionListDTO> getList() {
        return list;
    }

    public void setList(List<PensionListDTO> list) {
        this.list = list;
    }
}
