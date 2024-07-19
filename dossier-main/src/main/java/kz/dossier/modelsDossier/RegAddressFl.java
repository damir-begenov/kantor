package kz.dossier.modelsDossier;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;


@Entity
@Table(name = "fl_reg_address", schema = "initial_data")
public class RegAddressFl {
        @Column(name = "ID")
        private String id;
        private String fl_id;
        @Id
        @Column(name = "IIN")
        private String iin;
        private String registration_type;
        private String country;
        private String region;
        private String district;
        private String city;
        private String street;
        private String building;
        private String korpus;
        private String apartment_number;
        private String reg_date;
        private String reg_end_date;
        private String reg_reason;
        private String is_active;

        public String getId() {
                return id;
        }

        public void setId(String id) {
                this.id = id;
        }

        public String getFl_id() {
                return fl_id;
        }

        public void setFl_id(String fl_id) {
                this.fl_id = fl_id;
        }

        public String getIin() {
                return iin;
        }

        public void setIin(String iin) {
                this.iin = iin;
        }

        public String getRegistration_type() {
                return registration_type;
        }

        public void setRegistration_type(String registration_type) {
                this.registration_type = registration_type;
        }

        public String getCountry() {
                return country;
        }

        public void setCountry(String country) {
                this.country = country;
        }

        public String getRegion() {
                return region;
        }

        public void setRegion(String region) {
                this.region = region;
        }

        public String getDistrict() {
                return district;
        }

        public void setDistrict(String district) {
                this.district = district;
        }

        public String getCity() {
                return city;
        }

        public void setCity(String city) {
                this.city = city;
        }

        public String getStreet() {
                return street;
        }

        public void setStreet(String street) {
                this.street = street;
        }

        public String getBuilding() {
                return building;
        }

        public void setBuilding(String building) {
                this.building = building;
        }

        public String getKorpus() {
                return korpus;
        }

        public void setKorpus(String korpus) {
                this.korpus = korpus;
        }

        public String getApartment_number() {
                return apartment_number;
        }

        public void setApartment_number(String apartment_number) {
                this.apartment_number = apartment_number;
        }

        public String getReg_date() {
                return reg_date;
        }

        public void setReg_date(String reg_date) {
                this.reg_date = reg_date;
        }

        public String getReg_end_date() {
                return reg_end_date;
        }

        public void setReg_end_date(String reg_end_date) {
                this.reg_end_date = reg_end_date;
        }

        public String getReg_reason() {
                return reg_reason;
        }

        public void setReg_reason(String reg_reason) {
                this.reg_reason = reg_reason;
        }

        public String getIs_active() {
                return is_active;
        }

        public void setIs_active(String is_active) {
                this.is_active = is_active;
        }
}
