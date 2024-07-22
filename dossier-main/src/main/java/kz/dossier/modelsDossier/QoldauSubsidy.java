package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.sql.Timestamp;

@Entity
@Table(name = "qoldau_subsidy", schema = "imp_kfm_fl")
public class QoldauSubsidy {

    @Id
    @Nullable
    private Long id;

    @Nullable
    @Column(name = "name")
    private String name;
    @Nullable

    @Column(name = "region")
    private String region;
    @Nullable

    @Column(name = "date_of_application")
    private Timestamp dateOfApplication;
    @Nullable

    @Column(name = "applicant_name")
    private String applicantName;
    @Nullable

    @Column(name = "application_status")
    private String applicationStatus;
    @Nullable

    @Column(name = "date_of_acceptance")
    private Timestamp dateOfAcceptance;
    @Nullable

    @Column(name = "rejection_date")
    private Timestamp rejectionDate;
    @Nullable

    @Column(name = "subsidies_amount")
    private Double subsidiesAmount;
    @Nullable

    @Column(name = "rejection_reason")
    private String rejectionReason;
    @Nullable

    @Column(name = "note_to_decision")
    private String noteToDecision;
    @Nullable

    @Column(name = "ip_submission_app")
    private String ipSubmissionApp;
    @Nullable

    @Column(name = "ip_withdrawal_app")
    private String ipWithdrawalApp;
    @Nullable

    @Column(name = "ip_accentance_app")
    private String ipAcceptanceApp;
    @Nullable

    @Column(name = "ip_rejection_app")
    private String ipRejectionApp;
    @Nullable

    @Column(name = "bin")
    private String bin;

    @Nullable
    public Long getId() {
        return id;
    }

    public void setId(@Nullable Long id) {
        this.id = id;
    }

    @Nullable
    public String getName() {
        return name;
    }

    public void setName(@Nullable String name) {
        this.name = name;
    }

    @Nullable
    public String getRegion() {
        return region;
    }

    public void setRegion(@Nullable String region) {
        this.region = region;
    }

    @Nullable
    public Timestamp getDateOfApplication() {
        return dateOfApplication;
    }

    public void setDateOfApplication(@Nullable Timestamp dateOfApplication) {
        this.dateOfApplication = dateOfApplication;
    }

    @Nullable
    public String getApplicantName() {
        return applicantName;
    }

    public void setApplicantName(@Nullable String applicantName) {
        this.applicantName = applicantName;
    }

    @Nullable
    public String getApplicationStatus() {
        return applicationStatus;
    }

    public void setApplicationStatus(@Nullable String applicationStatus) {
        this.applicationStatus = applicationStatus;
    }

    @Nullable
    public Timestamp getDateOfAcceptance() {
        return dateOfAcceptance;
    }

    public void setDateOfAcceptance(@Nullable Timestamp dateOfAcceptance) {
        this.dateOfAcceptance = dateOfAcceptance;
    }

    @Nullable
    public Timestamp getRejectionDate() {
        return rejectionDate;
    }

    public void setRejectionDate(@Nullable Timestamp rejectionDate) {
        this.rejectionDate = rejectionDate;
    }

    @Nullable
    public Double getSubsidiesAmount() {
        return subsidiesAmount;
    }

    public void setSubsidiesAmount(@Nullable Double subsidiesAmount) {
        this.subsidiesAmount = subsidiesAmount;
    }

    @Nullable
    public String getRejectionReason() {
        return rejectionReason;
    }

    public void setRejectionReason(@Nullable String rejectionReason) {
        this.rejectionReason = rejectionReason;
    }

    @Nullable
    public String getNoteToDecision() {
        return noteToDecision;
    }

    public void setNoteToDecision(@Nullable String noteToDecision) {
        this.noteToDecision = noteToDecision;
    }

    @Nullable
    public String getIpSubmissionApp() {
        return ipSubmissionApp;
    }

    public void setIpSubmissionApp(@Nullable String ipSubmissionApp) {
        this.ipSubmissionApp = ipSubmissionApp;
    }

    @Nullable
    public String getIpWithdrawalApp() {
        return ipWithdrawalApp;
    }

    public void setIpWithdrawalApp(@Nullable String ipWithdrawalApp) {
        this.ipWithdrawalApp = ipWithdrawalApp;
    }

    @Nullable
    public String getIpAcceptanceApp() {
        return ipAcceptanceApp;
    }

    public void setIpAcceptanceApp(@Nullable String ipAcceptanceApp) {
        this.ipAcceptanceApp = ipAcceptanceApp;
    }

    @Nullable
    public String getIpRejectionApp() {
        return ipRejectionApp;
    }

    public void setIpRejectionApp(@Nullable String ipRejectionApp) {
        this.ipRejectionApp = ipRejectionApp;
    }

    @Nullable
    public String getBin() {
        return bin;
    }

    public void setBin(@Nullable String bin) {
        this.bin = bin;
    }
}