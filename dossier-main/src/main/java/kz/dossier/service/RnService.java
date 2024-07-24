package kz.dossier.service;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import kz.dossier.dto.RnDTO;
import kz.dossier.modelsDossier.MvRnOld;
import kz.dossier.repositoryDossier.MvRnOldRepo;

@Service
public class RnService {
    @Autowired
    MvRnOldRepo mvRnOldRepo;

    public void getDetailedRnView(String cadastrial_number, String address) {
        List<MvRnOld> rns = mvRnOldRepo.getRowsByCadAndAddress(cadastrial_number, address);
        List<RnDTO> rnDTOs = new ArrayList<>();
        for (MvRnOld rn: rns) {
            RnDTO rnDTO = new RnDTO();

            rnDTO.setNameOfKind("");
            rnDTO.setCadastrialNumber(cadastrial_number);
            rnDTO.setRightOwner("");
            rnDTO.setAddress(address);
            rnDTO.setFloorness(rn.getFloor());
            rnDTO.setSumOfDeal(rn.getRegister_transaction_amount());
            rnDTO.setAllArea(rn.getArea_total());
            rnDTO.setLivingArea(rn.getArea_useful());
            rnDTO.setTypeOfDoc("");
            rnDTO.setDocumentNumber(rn.getRegister_emergence_rights_rus());
            rnDTO.setDate("");
            if (rn.getRegister_end_date() != null) {
                rnDTO.setStatusRn("Исторический");
            } else {
                rnDTO.setStatusRn("Текущий");
            }
            rnDTO.setInOfOwner(rn.getOwner_iin_bin());
            rnDTO.setOwnerName(rn.getOwner_full_name());
            rnDTO.setDateOfRegistration(rn.getRegister_reg_date() + " - " + rn.getRegister_end_date());

            rnDTOs.add(rnDTO);
        }
    }
}
