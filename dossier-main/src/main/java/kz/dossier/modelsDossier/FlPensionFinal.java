package kz.dossier.modelsDossier;

import java.util.List;
import java.util.Map;

public class FlPensionFinal {
    private String CompanyBin;
    private List<String> years;
    private List<Map<String,Object>> nakoplenya;
    private List<Map<String,Object>> flPensionMinis;
    private Integer AmountOfEmp;
    private Integer year;

    public Integer getYear() {
        return year;
    }

    public void setYear(Integer year) {
        this.year = year;
    }

    public Integer getAmountOfEmp() {
        return AmountOfEmp;
    }

    public void setAmountOfEmp(Integer amountOfEmp) {
        AmountOfEmp = amountOfEmp;
    }



    public String getCompanyBin() {
        return CompanyBin;
    }

    public void setCompanyBin(String companyBin) {
        CompanyBin = companyBin;
    }

    public List<String> getYears() {
        return years;
    }

    public void setYears(List<String> years) {
        this.years = years;
    }

    public List<Map<String, Object>> getNakoplenya() {
        return nakoplenya;
    }

    public void setNakoplenya(List<Map<String, Object>> nakoplenya) {
        this.nakoplenya = nakoplenya;
    }

    public List<Map<String, Object>> getFlPensionMinis() {
        return flPensionMinis;
    }

    public void setFlPensionMinis(List<Map<String, Object>> flPensionMinis) {
        this.flPensionMinis = flPensionMinis;
    }
}
