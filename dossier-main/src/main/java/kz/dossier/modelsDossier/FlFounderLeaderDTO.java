package kz.dossier.modelsDossier;

import lombok.Data;

import java.util.List;

@Data
public class FlFounderLeaderDTO {
    private List<MvUlFounderFl> mvUlFounderFlList;
    private List<MvUlLeaderEntity> mvUlLeaderEntities;
    private int quantity;
}
