package kz.dossier.modelsDossier;

import org.jetbrains.annotations.Nullable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "tax_out", schema = "ser")
public class TaxOutEntity {
        @Id
        @Column(name = "Номер платежного документа")
        @Nullable
        private String paymentDocumentNumber;

        @Column(name = "бин")
        @Nullable
        private String bin;

        @Column(name = "Код бюджетной классификации")
        @Nullable
        private String budgetClassificationCode;

        @Column(name = "Дата зачисления в бюджет")
        @Nullable
        private String dateOfBudgetEnrollment;

        @Column(name = "Наименование КБК")
        @Nullable
        private String kbkName;

        @Column(name = "Код ОГД")
        @Nullable
        private String ogdCode;

        @Column(name = "Наименование ОГД")
        @Nullable
        private String ogdName;

        @Column(name = "Сумма платежа")
        @Nullable
        private String paymentAmount;

        @Column(name = "Тип платежа")
        @Nullable
        private String paymentType;

        @Column(name = "Тип проводки")
        @Nullable
        private String transactionType;

        @Column(name = "Дата списания со счета")
        @Nullable
        private String dateOfAccountDebit;

        @Nullable
        public String getPaymentDocumentNumber() {
                return paymentDocumentNumber;
        }

        public void setPaymentDocumentNumber(@Nullable String paymentDocumentNumber) {
                this.paymentDocumentNumber = paymentDocumentNumber;
        }

        @Nullable
        public String getBin() {
                return bin;
        }

        public void setBin(@Nullable String bin) {
                this.bin = bin;
        }

        @Nullable
        public String getBudgetClassificationCode() {
                return budgetClassificationCode;
        }

        public void setBudgetClassificationCode(@Nullable String budgetClassificationCode) {
                this.budgetClassificationCode = budgetClassificationCode;
        }

        @Nullable
        public String getDateOfBudgetEnrollment() {
                return dateOfBudgetEnrollment;
        }

        public void setDateOfBudgetEnrollment(@Nullable String dateOfBudgetEnrollment) {
                this.dateOfBudgetEnrollment = dateOfBudgetEnrollment;
        }

        @Nullable
        public String getKbkName() {
                return kbkName;
        }

        public void setKbkName(@Nullable String kbkName) {
                this.kbkName = kbkName;
        }

        @Nullable
        public String getOgdCode() {
                return ogdCode;
        }

        public void setOgdCode(@Nullable String ogdCode) {
                this.ogdCode = ogdCode;
        }

        @Nullable
        public String getOgdName() {
                return ogdName;
        }

        public void setOgdName(@Nullable String ogdName) {
                this.ogdName = ogdName;
        }

        @Nullable
        public String getPaymentAmount() {
                return paymentAmount;
        }

        public void setPaymentAmount(@Nullable String paymentAmount) {
                this.paymentAmount = paymentAmount;
        }

        @Nullable
        public String getPaymentType() {
                return paymentType;
        }

        public void setPaymentType(@Nullable String paymentType) {
                this.paymentType = paymentType;
        }

        @Nullable
        public String getTransactionType() {
                return transactionType;
        }

        public void setTransactionType(@Nullable String transactionType) {
                this.transactionType = transactionType;
        }

        @Nullable
        public String getDateOfAccountDebit() {
                return dateOfAccountDebit;
        }

        public void setDateOfAccountDebit(@Nullable String dateOfAccountDebit) {
                this.dateOfAccountDebit = dateOfAccountDebit;
        }
}