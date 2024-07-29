package kz.dossier.controller;

import kz.dossier.neo4j.entity.DTO.doubleReturn;
import kz.dossier.security.models.User;
import kz.dossier.security.repository.LogRepo;
import kz.dossier.security.repository.RoleRepository;
import kz.dossier.security.repository.UserRepository;
import kz.dossier.security.repository.UserRolesRepo;
import kz.dossier.security.services.LogsService;
import kz.dossier.security.services.UserDetailsServiceImpl;
import kz.dossier.service.StatisticService;
import kz.dossier.service.StoreLogs;
import kz.dossier.service.ZAGSDBService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.security.Principal;
import java.util.List;

@CrossOrigin(origins = "*", maxAge = 3000)
@RestController
@RequestMapping("/api/pandora/itapz")
public class ZagsDBController {
    @Autowired

    UserDetailsServiceImpl userDetailsService;    @Autowired

    StatisticService statisticService;    @Autowired

    LogRepo logRepo;

    @Autowired
    UserRepository userRepository;


    @Autowired
    LogsService logsService;
    @Autowired
    RoleRepository roleRepository;
    @Autowired
    UserRolesRepo userRolesRepo;

    @Autowired
    ZAGSDBService zagsdbService;
    @Autowired
    StoreLogs storeLogs;
    @GetMapping("/fltree")
    public doubleReturn getFlTree(Principal principal, @RequestParam String person, @RequestParam List<String> relations, @RequestParam int depth, @RequestParam int limit,
                                  @RequestParam(required = false) String orderNum, @RequestParam(required = false) String approvement_type,
                                  @RequestParam(required = false) String caseNum,
                                  @RequestParam(required = false) String orderDate,
                                  @RequestParam(required = false) String articleName,
                                  @RequestParam(required = false) String checkingName,
                                  @RequestParam(required = false) String otherReasons,
                                  @RequestParam(required = false) String organName,
                                  @RequestParam(required = false) String sphereName,
                                  @RequestParam(required = false) String tematikName,
                                  @RequestParam(required = false) String rukName) {
        User user = userDetailsService.loadUserByUsernamek(principal);
        relations = zagsdbService.filterRelations(user.getId(), relations);
        if (storeLogs.store(user, person, relations, depth, limit, orderNum, approvement_type,
                caseNum,
                orderDate,
                articleName,
                checkingName,
                otherReasons,
                organName,
                sphereName,
                tematikName,
                rukName)) {
            return zagsdbService.getPersonTree(user.getId(), person, depth, limit, relations);
        } else {
            return null;
        }
    }
    @GetMapping("/flFIOtree")
    public doubleReturn getFlTree(@RequestParam String lastName1,@RequestParam String firstName1,@RequestParam String fatherName1, @RequestParam List<String> relations, @RequestParam int depth, @RequestParam int limit, @RequestParam(required = false) String orderNum,
                                  @RequestParam(required = false) String caseNum, @RequestParam(required = false) String approvement_type,
                                  @RequestParam(required = false) String orderDate,
                                  @RequestParam(required = false) String articleName,
                                  @RequestParam(required = false) String checkingName,
                                  @RequestParam(required = false) String otherReasons,
                                  @RequestParam(required = false) String organName,
                                  @RequestParam(required = false) String sphereName,
                                  @RequestParam(required = false) String tematikName,
                                  @RequestParam(required = false) String rukName,Principal principal,
                                  @RequestParam(required = false) boolean check1) {
        User user = userDetailsService.loadUserByUsernamek(principal);
        relations = zagsdbService.filterRelations(user.getId(), relations);
        if (storeLogs.store(user, lastName1, firstName1, fatherName1, relations, depth, limit, orderNum, approvement_type,
                caseNum,
                orderDate,
                articleName,
                checkingName,
                otherReasons,
                organName,
                sphereName,
                tematikName,
                rukName)) {
            if (fatherName1.equals(".*") || fatherName1.equals(".*.*") ){
                return zagsdbService.getPersonByFIO_withoutO(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(), depth, limit, relations);
            }
            if (fatherName1.contains(".*") ){
                return zagsdbService.getPersonByFIO_(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(), depth, limit, relations);
            }
            if(fatherName1.equals("")  & check1 == true){
                return zagsdbService.getPersonByFIO_Accurate(user.getId(),lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),depth,limit,relations);
            }
            if(!fatherName1.equals("")  & check1 == true){
                return zagsdbService.getPersonByFIO_Accurate(user.getId(),lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),depth,limit,relations);
            }
            return zagsdbService.getPersonByFIO_(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(), depth, limit, relations);
        } else {
            return null;
        }
    }


    @GetMapping("/shortestpaths")
    public doubleReturn getShortestPaths(@RequestParam String person, @RequestParam String person2, @RequestParam List<String> relations, @RequestParam(required = false) String orderNum,
                                         @RequestParam(required = false) String caseNum, @RequestParam(required = false) String approvement_type,
                                         @RequestParam(required = false) String orderDate,
                                         @RequestParam(required = false) String articleName,
                                         @RequestParam(required = false) String checkingName,
                                         @RequestParam(required = false) String otherReasons,
                                         @RequestParam(required = false) String organName,
                                         @RequestParam(required = false) String sphereName,
                                         @RequestParam(required = false) String tematikName,
                                         @RequestParam(required = false) String rukName,Principal principal) {
        User user = userDetailsService.loadUserByUsernamek(principal);
        relations = zagsdbService.filterRelations(user.getId(), relations);
        if (storeLogs.store(user, person, person2, relations, orderNum, approvement_type,
                caseNum,
                orderDate,
                articleName,
                checkingName,
                otherReasons,
                organName,
                sphereName,
                tematikName,
                rukName)) {
            return zagsdbService.getShortestPaths(user.getId(), person, person2, relations);
        } else {
            return null;
        }
    }
    @GetMapping("/shortestpathsByFIO")
    public doubleReturn getShortestPathsByFIO(@RequestParam String lastName1,@RequestParam String firstName1,@RequestParam String fatherName1,@RequestParam String lastName2,@RequestParam String firstName2,@RequestParam String fatherName2, @RequestParam List<String> relations, @RequestParam(required = false) String orderNum,
                                              @RequestParam(required = false) String caseNum, @RequestParam(required = false) String approvement_type,
                                              @RequestParam(required = false) String orderDate,
                                              @RequestParam(required = false) String articleName,
                                              @RequestParam(required = false) String checkingName,
                                              @RequestParam(required = false) String otherReasons,
                                              @RequestParam(required = false) String organName,
                                              @RequestParam(required = false) String sphereName,
                                              @RequestParam(required = false) String tematikName,
                                              @RequestParam(required = false) String rukName,Principal principal,
                                              @RequestParam(required = false) boolean check1,
                                              @RequestParam(required = false) boolean check2) {
        User user = userDetailsService.loadUserByUsernamek(principal);
        relations = zagsdbService.filterRelations(user.getId(), relations);
        if (storeLogs.store(lastName1, firstName1, fatherName1, lastName2, firstName2, fatherName2, relations, orderNum,
                caseNum, approvement_type,
                orderDate,
                articleName,
                checkingName,
                otherReasons,
                organName,
                sphereName,
                tematikName,
                rukName, user, check1, check2)) {

            if(check1 == true & check2 == true){
                return zagsdbService.getShortestPathWithFIOAccurate(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }
            if(check1 == true & check2 == false){
                return zagsdbService.getShortestPathWithFIOAccurate12(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }
            if(check1 == false & check2 == true){
                return zagsdbService.getShortestPathWithFIOAccurate21(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }
            if(fatherName1 == "" & fatherName2 == ""){
                return zagsdbService.getShortestPathWithFIObezbez(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }
            if(fatherName1 == "" & fatherName2 != ""){
                return zagsdbService.getShortestPathWithFIObezs(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }
            if(fatherName1 != "" & fatherName2 == ""){
                return zagsdbService.getShortestPathWithFIOsbez(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);
            }

            return zagsdbService.getShortestPathWithFIO(user.getId(), lastName1.toUpperCase(),firstName1.toUpperCase(),fatherName1.toUpperCase(),lastName2.toUpperCase(),firstName2.toUpperCase(),fatherName2.toUpperCase(), relations);        } else {
            return null;
        }

    }
    @GetMapping("/shortopen")
    public doubleReturn getShortOpen(@RequestParam Long id, @RequestParam List<String> relations, @RequestParam int limit, Principal principal) {
        User user = userDetailsService.loadUserByUsernamek(principal);
        relations = zagsdbService.filterRelations(user.getId(), relations);
        doubleReturn resul = zagsdbService.shortOpen(user.getId(), id, relations, limit);
        return resul;
    }
}
