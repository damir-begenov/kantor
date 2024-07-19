package kz.dossier.service;

import kz.dossier.modelsDossier.PhotoDb;
import kz.dossier.repositoryDossier.NewPhotoRepo;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class PhotoService {
    @Autowired
    NewPhotoRepo newPhotoRepo;

    public PhotoDb getPhotoByIIN(String iin){
        return newPhotoRepo.findById(iin).orElse(null);
    }
}
