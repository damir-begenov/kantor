package kz.dossier.modelsDossier;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table(name = "снятые", schema = "transport")
@Data
public class AutoSnyatie {
    @Id
    @Column(name = "ВИН код ТС")
    private String vin;
    @Column(name = "Марка, модель, модификация  ТС")
    private String model;
    @Column(name = "Государственный номерной знак (ГР")
    private String nomer;
    @Column(name = "Дата первичной регистрации ТС")
    private String date_registration;
    @Column(name = "Серия и номер свидетельства о реги")
    private String series;
    @Column(name = "Категория ТС")
    private String category;
    @Column(name = "Объем двигателя ТС (см. куб.)")
    private String obem;
    @Column(name = "Масса без нагрузки (кг.)")
    private String massa;
    @Column(name = "Разрешенная максимальная масса (к")
    private String max_mass;
    @Column(name = "Количество посадочных мест")
    private String seats;
    @Column(name = "Категория правообладателя")
    private String license_category;
    @Column(name = "БИН субъекта")
    private String bin;


}
