package kz.dossier.controller;

import kz.dossier.security.models.log;
import kz.dossier.security.repository.LogRepo;
import lombok.AllArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.security.Principal;
import java.util.List;

@RestController
@AllArgsConstructor
@CrossOrigin(origins = "*", maxAge = 3600)
@RequestMapping("/api/pandora/main")
public class LogController {
    @Autowired
    LogRepo logRepo;

    @GetMapping("/admin/searchuserlogs")
    public List<log> searchByIdAndId(@RequestParam String value, @RequestParam String username) {
        List<log> result = logRepo.findByUsernameAndInput(username, value);
        return result;
    }

    @GetMapping("/logs")
    public List<log> searchByIdAndId(Principal principal) {
        String username = principal.getName();
        List<log> result = logRepo.findByUsername(username);
        return result;
    }


}
