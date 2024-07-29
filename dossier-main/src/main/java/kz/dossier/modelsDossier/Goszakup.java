package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

import java.sql.Date;

@Entity
@Table(name = "gos_zakup", schema = "imp_kfm_db")
public class Goszakup {
    @Id
    private Long id;
    @Column(name = "parent_id")
    private String parentId;

    @Column(name = "root_id")
    private String rootId;

    @Column(name = "trd_buy_id")
    private String trdBuyId;

    @Column(name = "trd_buy_number_anno")
    private String trdBuyNumberAnno;

    @Column(name = "trd_buy_name_ru")
    private String trdBuyNameRu;

    @Column(name = "trd_buy_name_kz")
    private String trdBuyNameKz;

    @Column(name = "ref_amendm_agreem_justif_id")
    private String refAmendmAgreemJustifId;

    @Column(name = "ref_contract_status_id")
    private String refContractStatusId;

    @Column(name = "deleted")
    private String deleted;

    @Column(name = "crdate")
    private String crdate;

    @Column(name = "crdatee")
    private String crdatee;

    @Column(name = "last_update_date")
    private String lastUpdateDate;

    @Column(name = "supplier_id")
    private String supplierId;

    @Column(name = "supplier_bin")
    private String supplierBin;

    @Column(name = "supplier_bik")
    private String supplierBik;

    @Column(name = "supplier_iik")
    private String supplierIik;

    @Column(name = "supplier_bank_name_kz")
    private String supplierBankNameKz;

    @Column(name = "supplier_bank_name_ru")
    private String supplierBankNameRu;

    @Column(name = "supplier_legal_address")
    private String supplierLegalAddress;

    @Column(name = "supplier_bill_id")
    private String supplierBillId;

    @Column(name = "contract_number")
    private String contractNumber;

    @Column(name = "sign_reason_doc_name")
    private String signReasonDocName;

    @Column(name = "sign_reason_doc_date")
    private String signReasonDocDate;

    @Column(name = "trd_buy_itogi_date_public")
    private String trdBuyItogiDatePublic;

    @Column(name = "customer_id")
    private String customerId;

    @Column(name = "customer_bin")
    private String customerBin;

    @Column(name = "customer_bik")
    private String customerBik;

    @Column(name = "customer_iik")
    private String customerIik;

    @Column(name = "customer_bank_name_kz")
    private String customerBankNameKz;

    @Column(name = "customer_bank_name_ru")
    private String customerBankNameRu;

    @Column(name = "customer_legal_address")
    private String customerLegalAddress;

    @Column(name = "customer_bill_id")
    private String customerBillId;

    @Column(name = "contract_number_sys")
    private String contractNumberSys;

    @Column(name = "payments_terms_ru")
    private String paymentsTermsRu;

    @Column(name = "payments_terms_kz")
    private String paymentsTermsKz;

    @Column(name = "ref_subject_type_id")
    private String refSubjectTypeId;

    @Column(name = "ref_subject_types_id")
    private String refSubjectTypesId;

    @Column(name = "is_gu")
    private String isGu;

    @Column(name = "fin_year")
    private String finYear;

    @Column(name = "ref_contract_agr_form_id")
    private String refContractAgrFormId;

    @Column(name = "ref_contract_year_type_id")
    private String refContractYearTypeId;

    @Column(name = "ref_finsource_id")
    private String refFinsourceId;

    @Column(name = "ref_currency_code")
    private String refCurrencyCode;

    @Column(name = "exchange_rate")
    private String exchangeRate;

    @Column(name = "contract_sum")
    private String contractSum;

    @Column(name = "contract_sum_wnds")
    private double contractSumWnds;

    @Column(name = "sign_date")
    private Date signDate;

    @Column(name = "sign_datee")
    private String signDatee;

    @Column(name = "ec_end_date")
    private String ecEndDate;

    @Column(name = "plan_exec_date")
    private String planExecDate;

    @Column(name = "fakt_exec_date")
    private String faktExecDate;

    @Column(name = "fakt_sum")
    private String faktSum;

    @Column(name = "fakt_sum_wnds")
    private String faktSumWnds;

    @Column(name = "contract_end_date")
    private String contractEndDate;

    @Column(name = "ref_contract_cancel_id")
    private String refContractCancelId;

    @Column(name = "ref_contract_type_id")
    private String refContractTypeId;

    @Column(name = "description_kz")
    private String descriptionKz;

    @Column(name = "description_ru")
    private String descriptionRu;

    @Column(name = "fakt_trade_methods_id")
    private String faktTradeMethodsId;

    @Column(name = "ec_customer_approve")
    private String ecCustomerApprove;

    @Column(name = "ec_supplier_approve")
    private String ecSupplierApprove;

    @Column(name = "contract_ms")
    private String contractMs;

    @Column(name = "treasure_req_num")
    private String treasureReqNum;

    @Column(name = "treasure_req_date")
    private String treasureReqDate;

    @Column(name = "treasure_not_num")
    private String treasureNotNum;

    @Column(name = "treasure_not_date")
    private String treasureNotDate;

    @Column(name = "system_id")
    private String systemId;

    @Column(name = "index_date")
    private String indexDate;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getParentId() {
        return parentId;
    }

    public void setParentId(String parentId) {
        this.parentId = parentId;
    }

    public String getRootId() {
        return rootId;
    }

    public void setRootId(String rootId) {
        this.rootId = rootId;
    }

    public String getTrdBuyId() {
        return trdBuyId;
    }

    public void setTrdBuyId(String trdBuyId) {
        this.trdBuyId = trdBuyId;
    }

    public String getTrdBuyNumberAnno() {
        return trdBuyNumberAnno;
    }

    public void setTrdBuyNumberAnno(String trdBuyNumberAnno) {
        this.trdBuyNumberAnno = trdBuyNumberAnno;
    }

    public String getTrdBuyNameRu() {
        return trdBuyNameRu;
    }

    public void setTrdBuyNameRu(String trdBuyNameRu) {
        this.trdBuyNameRu = trdBuyNameRu;
    }

    public String getTrdBuyNameKz() {
        return trdBuyNameKz;
    }

    public void setTrdBuyNameKz(String trdBuyNameKz) {
        this.trdBuyNameKz = trdBuyNameKz;
    }

    public String getRefAmendmAgreemJustifId() {
        return refAmendmAgreemJustifId;
    }

    public void setRefAmendmAgreemJustifId(String refAmendmAgreemJustifId) {
        this.refAmendmAgreemJustifId = refAmendmAgreemJustifId;
    }

    public String getRefContractStatusId() {
        return refContractStatusId;
    }

    public void setRefContractStatusId(String refContractStatusId) {
        this.refContractStatusId = refContractStatusId;
    }

    public String getDeleted() {
        return deleted;
    }

    public void setDeleted(String deleted) {
        this.deleted = deleted;
    }

    public String getCrdate() {
        return crdate;
    }

    public void setCrdate(String crdate) {
        this.crdate = crdate;
    }

    public String getCrdatee() {
        return crdatee;
    }

    public void setCrdatee(String crdatee) {
        this.crdatee = crdatee;
    }

    public String getLastUpdateDate() {
        return lastUpdateDate;
    }

    public void setLastUpdateDate(String lastUpdateDate) {
        this.lastUpdateDate = lastUpdateDate;
    }

    public String getSupplierId() {
        return supplierId;
    }

    public void setSupplierId(String supplierId) {
        this.supplierId = supplierId;
    }

    public String getSupplierBin() {
        return supplierBin;
    }

    public void setSupplierBin(String supplierBin) {
        this.supplierBin = supplierBin;
    }

    public String getSupplierBik() {
        return supplierBik;
    }

    public void setSupplierBik(String supplierBik) {
        this.supplierBik = supplierBik;
    }

    public String getSupplierIik() {
        return supplierIik;
    }

    public void setSupplierIik(String supplierIik) {
        this.supplierIik = supplierIik;
    }

    public String getSupplierBankNameKz() {
        return supplierBankNameKz;
    }

    public void setSupplierBankNameKz(String supplierBankNameKz) {
        this.supplierBankNameKz = supplierBankNameKz;
    }

    public String getSupplierBankNameRu() {
        return supplierBankNameRu;
    }

    public void setSupplierBankNameRu(String supplierBankNameRu) {
        this.supplierBankNameRu = supplierBankNameRu;
    }

    public String getSupplierLegalAddress() {
        return supplierLegalAddress;
    }

    public void setSupplierLegalAddress(String supplierLegalAddress) {
        this.supplierLegalAddress = supplierLegalAddress;
    }

    public String getSupplierBillId() {
        return supplierBillId;
    }

    public void setSupplierBillId(String supplierBillId) {
        this.supplierBillId = supplierBillId;
    }

    public String getContractNumber() {
        return contractNumber;
    }

    public void setContractNumber(String contractNumber) {
        this.contractNumber = contractNumber;
    }

    public String getSignReasonDocName() {
        return signReasonDocName;
    }

    public void setSignReasonDocName(String signReasonDocName) {
        this.signReasonDocName = signReasonDocName;
    }

    public String getSignReasonDocDate() {
        return signReasonDocDate;
    }

    public void setSignReasonDocDate(String signReasonDocDate) {
        this.signReasonDocDate = signReasonDocDate;
    }

    public String getTrdBuyItogiDatePublic() {
        return trdBuyItogiDatePublic;
    }

    public void setTrdBuyItogiDatePublic(String trdBuyItogiDatePublic) {
        this.trdBuyItogiDatePublic = trdBuyItogiDatePublic;
    }

    public String getCustomerId() {
        return customerId;
    }

    public void setCustomerId(String customerId) {
        this.customerId = customerId;
    }

    public String getCustomerBin() {
        return customerBin;
    }

    public void setCustomerBin(String customerBin) {
        this.customerBin = customerBin;
    }

    public String getCustomerBik() {
        return customerBik;
    }

    public void setCustomerBik(String customerBik) {
        this.customerBik = customerBik;
    }

    public String getCustomerIik() {
        return customerIik;
    }

    public void setCustomerIik(String customerIik) {
        this.customerIik = customerIik;
    }

    public String getCustomerBankNameKz() {
        return customerBankNameKz;
    }

    public void setCustomerBankNameKz(String customerBankNameKz) {
        this.customerBankNameKz = customerBankNameKz;
    }

    public String getCustomerBankNameRu() {
        return customerBankNameRu;
    }

    public void setCustomerBankNameRu(String customerBankNameRu) {
        this.customerBankNameRu = customerBankNameRu;
    }

    public String getCustomerLegalAddress() {
        return customerLegalAddress;
    }

    public void setCustomerLegalAddress(String customerLegalAddress) {
        this.customerLegalAddress = customerLegalAddress;
    }

    public String getCustomerBillId() {
        return customerBillId;
    }

    public void setCustomerBillId(String customerBillId) {
        this.customerBillId = customerBillId;
    }

    public String getContractNumberSys() {
        return contractNumberSys;
    }

    public void setContractNumberSys(String contractNumberSys) {
        this.contractNumberSys = contractNumberSys;
    }

    public String getPaymentsTermsRu() {
        return paymentsTermsRu;
    }

    public void setPaymentsTermsRu(String paymentsTermsRu) {
        this.paymentsTermsRu = paymentsTermsRu;
    }

    public String getPaymentsTermsKz() {
        return paymentsTermsKz;
    }

    public void setPaymentsTermsKz(String paymentsTermsKz) {
        this.paymentsTermsKz = paymentsTermsKz;
    }

    public String getRefSubjectTypeId() {
        return refSubjectTypeId;
    }

    public void setRefSubjectTypeId(String refSubjectTypeId) {
        this.refSubjectTypeId = refSubjectTypeId;
    }

    public String getRefSubjectTypesId() {
        return refSubjectTypesId;
    }

    public void setRefSubjectTypesId(String refSubjectTypesId) {
        this.refSubjectTypesId = refSubjectTypesId;
    }

    public String getIsGu() {
        return isGu;
    }

    public void setIsGu(String isGu) {
        this.isGu = isGu;
    }

    public String getFinYear() {
        return finYear;
    }

    public void setFinYear(String finYear) {
        this.finYear = finYear;
    }

    public String getRefContractAgrFormId() {
        return refContractAgrFormId;
    }

    public void setRefContractAgrFormId(String refContractAgrFormId) {
        this.refContractAgrFormId = refContractAgrFormId;
    }

    public String getRefContractYearTypeId() {
        return refContractYearTypeId;
    }

    public void setRefContractYearTypeId(String refContractYearTypeId) {
        this.refContractYearTypeId = refContractYearTypeId;
    }

    public String getRefFinsourceId() {
        return refFinsourceId;
    }

    public void setRefFinsourceId(String refFinsourceId) {
        this.refFinsourceId = refFinsourceId;
    }

    public String getRefCurrencyCode() {
        return refCurrencyCode;
    }

    public void setRefCurrencyCode(String refCurrencyCode) {
        this.refCurrencyCode = refCurrencyCode;
    }

    public String getExchangeRate() {
        return exchangeRate;
    }

    public void setExchangeRate(String exchangeRate) {
        this.exchangeRate = exchangeRate;
    }

    public String getContractSum() {
        return contractSum;
    }

    public void setContractSum(String contractSum) {
        this.contractSum = contractSum;
    }

    public double getContractSumWnds() {
        return contractSumWnds;
    }

    public void setContractSumWnds(double contractSumWnds) {
        this.contractSumWnds = contractSumWnds;
    }

    public Date getSignDate() {
        return signDate;
    }

    public void setSignDate(Date signDate) {
        this.signDate = signDate;
    }

    public String getSignDatee() {
        return signDatee;
    }

    public void setSignDatee(String signDatee) {
        this.signDatee = signDatee;
    }

    public String getEcEndDate() {
        return ecEndDate;
    }

    public void setEcEndDate(String ecEndDate) {
        this.ecEndDate = ecEndDate;
    }

    public String getPlanExecDate() {
        return planExecDate;
    }

    public void setPlanExecDate(String planExecDate) {
        this.planExecDate = planExecDate;
    }

    public String getFaktExecDate() {
        return faktExecDate;
    }

    public void setFaktExecDate(String faktExecDate) {
        this.faktExecDate = faktExecDate;
    }

    public String getFaktSum() {
        return faktSum;
    }

    public void setFaktSum(String faktSum) {
        this.faktSum = faktSum;
    }

    public String getFaktSumWnds() {
        return faktSumWnds;
    }

    public void setFaktSumWnds(String faktSumWnds) {
        this.faktSumWnds = faktSumWnds;
    }

    public String getContractEndDate() {
        return contractEndDate;
    }

    public void setContractEndDate(String contractEndDate) {
        this.contractEndDate = contractEndDate;
    }

    public String getRefContractCancelId() {
        return refContractCancelId;
    }

    public void setRefContractCancelId(String refContractCancelId) {
        this.refContractCancelId = refContractCancelId;
    }

    public String getRefContractTypeId() {
        return refContractTypeId;
    }

    public void setRefContractTypeId(String refContractTypeId) {
        this.refContractTypeId = refContractTypeId;
    }

    public String getDescriptionKz() {
        return descriptionKz;
    }

    public void setDescriptionKz(String descriptionKz) {
        this.descriptionKz = descriptionKz;
    }

    public String getDescriptionRu() {
        return descriptionRu;
    }

    public void setDescriptionRu(String descriptionRu) {
        this.descriptionRu = descriptionRu;
    }

    public String getFaktTradeMethodsId() {
        return faktTradeMethodsId;
    }

    public void setFaktTradeMethodsId(String faktTradeMethodsId) {
        this.faktTradeMethodsId = faktTradeMethodsId;
    }

    public String getEcCustomerApprove() {
        return ecCustomerApprove;
    }

    public void setEcCustomerApprove(String ecCustomerApprove) {
        this.ecCustomerApprove = ecCustomerApprove;
    }

    public String getEcSupplierApprove() {
        return ecSupplierApprove;
    }

    public void setEcSupplierApprove(String ecSupplierApprove) {
        this.ecSupplierApprove = ecSupplierApprove;
    }

    public String getContractMs() {
        return contractMs;
    }

    public void setContractMs(String contractMs) {
        this.contractMs = contractMs;
    }

    public String getTreasureReqNum() {
        return treasureReqNum;
    }

    public void setTreasureReqNum(String treasureReqNum) {
        this.treasureReqNum = treasureReqNum;
    }

    public String getTreasureReqDate() {
        return treasureReqDate;
    }

    public void setTreasureReqDate(String treasureReqDate) {
        this.treasureReqDate = treasureReqDate;
    }

    public String getTreasureNotNum() {
        return treasureNotNum;
    }

    public void setTreasureNotNum(String treasureNotNum) {
        this.treasureNotNum = treasureNotNum;
    }

    public String getTreasureNotDate() {
        return treasureNotDate;
    }

    public void setTreasureNotDate(String treasureNotDate) {
        this.treasureNotDate = treasureNotDate;
    }

    public String getSystemId() {
        return systemId;
    }

    public void setSystemId(String systemId) {
        this.systemId = systemId;
    }

    public String getIndexDate() {
        return indexDate;
    }

    public void setIndexDate(String indexDate) {
        this.indexDate = indexDate;
    }
}
