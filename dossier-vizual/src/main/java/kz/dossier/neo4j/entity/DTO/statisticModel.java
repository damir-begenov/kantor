package kz.dossier.neo4j.entity.DTO;

import kz.dossier.security.models.User;
import kz.dossier.security.models.log;

import java.time.LocalDateTime;
import java.util.List;

public class statisticModel {
    private User user;
    private String role;
    private LocalDateTime lastDate;
    private int allRequestNum;
    private int todayRequestNum;
    private List<log> logs;

    public String getRole() {
        return role;
    }
    public void setRole(String role) {
        this.role = role;
    }
    public User getUser() {
        return user;
    }
    public void setUser(User user) {
        this.user = user;
    }
    public LocalDateTime getDate() {
        return lastDate;
    }
    public void setDate(LocalDateTime lastDate) {
        this.lastDate = lastDate;
    }
    public int getAllRequsetNum() {
        return allRequestNum;
    }
    public void setAllRequsetNum(int allRequsetNum) {
        this.allRequestNum = allRequsetNum;
    }
    public int getTodayRequsetNum() {
        return todayRequestNum;
    }
    public void setTodayRequsetNum(int todayRequsetNum) {
        this.todayRequestNum = todayRequsetNum;
    }
    public List<log> getLogs() {
        return logs;
    }
    public void setLogs(List<log> logs) {
        this.logs = logs;
    }

 
    
}
