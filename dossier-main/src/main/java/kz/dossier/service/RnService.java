package kz.dossier.service;

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
        for (MvRnOld rn : rns) {
            System.out.println(rn.getOwner_iin_bin() + " " + rn.getRegister_reg_date() + " " + rn.getRegister_end_date());
        }
    }
}
