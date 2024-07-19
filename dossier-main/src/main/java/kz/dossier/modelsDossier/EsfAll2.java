package kz.dossier.modelsDossier;


import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.sql.Date;

@Entity
@Table(name = "esf_new_2", schema = "esf")
public class EsfAll2 {
    @Id
    private String INVOICE_ID;
    private String PRODUCT_NUMBER;
    private String DESCRIPTION;
    private String UNIT_CODE;
    private String UNIT_NOMENCLATURE;
    private Integer QUANTITY;
    private Integer UNIT_PRICE;
    private Integer EXCISE_RATE;
    private Integer EXCISE_AMOUNT;
    private Integer TURNOVER_SIZE;
    private Integer NDS_RATE;
    private Integer NDS_AMOUNT;
    private String APPNUM_UNION;
    private String TRUORIGINCODE;
    private String IIN_SELLER;
    private String IIN_SELLER_ROOT;
    private String STATUS_SELLER;
    private String SHARE_PARTICIPATION_SELLER;
    private String IIN_CUSTOMER;
    private String IIN_CUSTOMER_ROOT;
    private String STATUS_CUSTOMER;
    private String SHARE_PARTICIPATION_CUSTOMER;
    private String REGISTRATION_NUMBER;
    private String INVOICE_TYPE;
    private String INVOICE_NUMBER;
    private Date INVOICE_DATE;
    private Date TURNOVER_DATE;
    private Integer TOTALTURNOVERSIZE;
    private Integer TOTALNDSAMOUNT;
    private String KOGD;
    private String Quartile;

    public String getINVOICE_ID() {
        return INVOICE_ID;
    }

    public void setINVOICE_ID(String INVOICE_ID) {
        this.INVOICE_ID = INVOICE_ID;
    }

    public String getPRODUCT_NUMBER() {
        return PRODUCT_NUMBER;
    }

    public void setPRODUCT_NUMBER(String PRODUCT_NUMBER) {
        this.PRODUCT_NUMBER = PRODUCT_NUMBER;
    }

    public String getDESCRIPTION() {
        return DESCRIPTION;
    }

    public void setDESCRIPTION(String DESCRIPTION) {
        this.DESCRIPTION = DESCRIPTION;
    }

    public String getUNIT_CODE() {
        return UNIT_CODE;
    }

    public void setUNIT_CODE(String UNIT_CODE) {
        this.UNIT_CODE = UNIT_CODE;
    }

    public String getUNIT_NOMENCLATURE() {
        return UNIT_NOMENCLATURE;
    }

    public void setUNIT_NOMENCLATURE(String UNIT_NOMENCLATURE) {
        this.UNIT_NOMENCLATURE = UNIT_NOMENCLATURE;
    }

    public Integer getQUANTITY() {
        return QUANTITY;
    }

    public void setQUANTITY(Integer QUANTITY) {
        this.QUANTITY = QUANTITY;
    }

    public Integer getUNIT_PRICE() {
        return UNIT_PRICE;
    }

    public void setUNIT_PRICE(Integer UNIT_PRICE) {
        this.UNIT_PRICE = UNIT_PRICE;
    }

    public Integer getEXCISE_RATE() {
        return EXCISE_RATE;
    }

    public void setEXCISE_RATE(Integer EXCISE_RATE) {
        this.EXCISE_RATE = EXCISE_RATE;
    }

    public Integer getEXCISE_AMOUNT() {
        return EXCISE_AMOUNT;
    }

    public void setEXCISE_AMOUNT(Integer EXCISE_AMOUNT) {
        this.EXCISE_AMOUNT = EXCISE_AMOUNT;
    }

    public Integer getTURNOVER_SIZE() {
        return TURNOVER_SIZE;
    }

    public void setTURNOVER_SIZE(Integer TURNOVER_SIZE) {
        this.TURNOVER_SIZE = TURNOVER_SIZE;
    }

    public Integer getNDS_RATE() {
        return NDS_RATE;
    }

    public void setNDS_RATE(Integer NDS_RATE) {
        this.NDS_RATE = NDS_RATE;
    }

    public Integer getNDS_AMOUNT() {
        return NDS_AMOUNT;
    }

    public void setNDS_AMOUNT(Integer NDS_AMOUNT) {
        this.NDS_AMOUNT = NDS_AMOUNT;
    }

    public String getAPPNUM_UNION() {
        return APPNUM_UNION;
    }

    public void setAPPNUM_UNION(String APPNUM_UNION) {
        this.APPNUM_UNION = APPNUM_UNION;
    }

    public String getTRUORIGINCODE() {
        return TRUORIGINCODE;
    }

    public void setTRUORIGINCODE(String TRUORIGINCODE) {
        this.TRUORIGINCODE = TRUORIGINCODE;
    }

    public String getIIN_SELLER() {
        return IIN_SELLER;
    }

    public void setIIN_SELLER(String IIN_SELLER) {
        this.IIN_SELLER = IIN_SELLER;
    }

    public String getIIN_SELLER_ROOT() {
        return IIN_SELLER_ROOT;
    }

    public void setIIN_SELLER_ROOT(String IIN_SELLER_ROOT) {
        this.IIN_SELLER_ROOT = IIN_SELLER_ROOT;
    }

    public String getSTATUS_SELLER() {
        return STATUS_SELLER;
    }

    public void setSTATUS_SELLER(String STATUS_SELLER) {
        this.STATUS_SELLER = STATUS_SELLER;
    }

    public String getSHARE_PARTICIPATION_SELLER() {
        return SHARE_PARTICIPATION_SELLER;
    }

    public void setSHARE_PARTICIPATION_SELLER(String SHARE_PARTICIPATION_SELLER) {
        this.SHARE_PARTICIPATION_SELLER = SHARE_PARTICIPATION_SELLER;
    }

    public String getIIN_CUSTOMER() {
        return IIN_CUSTOMER;
    }

    public void setIIN_CUSTOMER(String IIN_CUSTOMER) {
        this.IIN_CUSTOMER = IIN_CUSTOMER;
    }

    public String getIIN_CUSTOMER_ROOT() {
        return IIN_CUSTOMER_ROOT;
    }

    public void setIIN_CUSTOMER_ROOT(String IIN_CUSTOMER_ROOT) {
        this.IIN_CUSTOMER_ROOT = IIN_CUSTOMER_ROOT;
    }

    public String getSTATUS_CUSTOMER() {
        return STATUS_CUSTOMER;
    }

    public void setSTATUS_CUSTOMER(String STATUS_CUSTOMER) {
        this.STATUS_CUSTOMER = STATUS_CUSTOMER;
    }

    public String getSHARE_PARTICIPATION_CUSTOMER() {
        return SHARE_PARTICIPATION_CUSTOMER;
    }

    public void setSHARE_PARTICIPATION_CUSTOMER(String SHARE_PARTICIPATION_CUSTOMER) {
        this.SHARE_PARTICIPATION_CUSTOMER = SHARE_PARTICIPATION_CUSTOMER;
    }

    public String getREGISTRATION_NUMBER() {
        return REGISTRATION_NUMBER;
    }

    public void setREGISTRATION_NUMBER(String REGISTRATION_NUMBER) {
        this.REGISTRATION_NUMBER = REGISTRATION_NUMBER;
    }

    public String getINVOICE_TYPE() {
        return INVOICE_TYPE;
    }

    public void setINVOICE_TYPE(String INVOICE_TYPE) {
        this.INVOICE_TYPE = INVOICE_TYPE;
    }

    public String getINVOICE_NUMBER() {
        return INVOICE_NUMBER;
    }

    public void setINVOICE_NUMBER(String INVOICE_NUMBER) {
        this.INVOICE_NUMBER = INVOICE_NUMBER;
    }

    public Date getINVOICE_DATE() {
        return INVOICE_DATE;
    }

    public void setINVOICE_DATE(Date INVOICE_DATE) {
        this.INVOICE_DATE = INVOICE_DATE;
    }

    public Date getTURNOVER_DATE() {
        return TURNOVER_DATE;
    }

    public void setTURNOVER_DATE(Date TURNOVER_DATE) {
        this.TURNOVER_DATE = TURNOVER_DATE;
    }

    public Integer getTOTALTURNOVERSIZE() {
        return TOTALTURNOVERSIZE;
    }

    public void setTOTALTURNOVERSIZE(Integer TOTALTURNOVERSIZE) {
        this.TOTALTURNOVERSIZE = TOTALTURNOVERSIZE;
    }

    public Integer getTOTALNDSAMOUNT() {
        return TOTALNDSAMOUNT;
    }

    public void setTOTALNDSAMOUNT(Integer TOTALNDSAMOUNT) {
        this.TOTALNDSAMOUNT = TOTALNDSAMOUNT;
    }

    public String getKOGD() {
        return KOGD;
    }

    public void setKOGD(String KOGD) {
        this.KOGD = KOGD;
    }

    public String getQuartile() {
        return Quartile;
    }

    public void setQuartile(String quartile) {
        Quartile = quartile;
    }
}
