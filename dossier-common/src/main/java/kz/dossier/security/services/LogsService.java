package kz.dossier.security.services;

import kz.dossier.security.models.log;
import kz.dossier.security.repository.LogRepo;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class LogsService {
    @Autowired
    LogRepo logRepo;

    public log SaveLog(log log){
       return   logRepo.save(log);
    }
}
