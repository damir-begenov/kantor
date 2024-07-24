package kz.dossier.tools;

import kz.dossier.modelsDossier.*;
import kz.dossier.modelsRisk.*;
import kz.dossier.repositoryDossier.FlPensionMiniRepo;
import kz.dossier.repositoryDossier.MvUlRepo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class DocxGenerator {
    @Autowired
    FlPensionMiniRepo flPensionMiniRepo;
    private MvUlRepo mvUlRepo;

    private XWPFDocument makeTableByProperties(XWPFDocument doc, XWPFTable table, String title, List<String> properties) {
        table.setWidth("100%");
        XWPFTableRow row3 = table.createRow();
        XWPFTableCell cell = row3.addNewTableCell();
        cell.setWidth("100%");
        cell.setColor("808080");
        XWPFParagraph paragraph1 = cell.addParagraph();
        paragraph1.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph1.createRun();
        run.setText(title);
        XWPFTableRow row = table.createRow();
        for (String prop : properties) {
            row.addNewTableCell().setText(prop);
        }
        return doc;
    }

    private void setMarginBetweenTables(XWPFDocument doc) {
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run2 = paragraph.createRun();
        run2.addBreak();  // Добавляем перенос строки
        run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
    }

    public void generateDoc(NodesFL result, ByteArrayOutputStream baos) throws IOException, InvalidFormatException {
        try (XWPFDocument doc = new XWPFDocument()) {
            try {
                if (result.getMvFls() != null || result.getMvFls().size() < 0) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения о физическом лице", Arrays.asList(
                            "Фото",
                            "ИИН",
                            "ФИО",
                            "Резидент",
                            "Национальность",
                            "Дата смерти"));
                    XWPFTableRow row1 = table.createRow();
                    XWPFTableCell cell1 = row1.createCell();
                    XWPFParagraph paragraph2 = cell1.addParagraph();

                    setCellPadding(cell1, 200, 200, 200, 200);
                    XWPFRun run1 = paragraph2.createRun();

                    byte[] imageBytes = result.getPhotoDbf().get(0).getPhoto();
                    ByteArrayInputStream imageStream = new ByteArrayInputStream(imageBytes);
                    int imageType = XWPFDocument.PICTURE_TYPE_PNG; // Change according to your image type (e.g., PICTURE_TYPE_JPEG)
                    run1.addPicture(imageStream, imageType, "image.png", Units.toEMU(75), Units.toEMU(100));

                    row1.addNewTableCell().setText(result.getMvFls().get(0).getIin());
                    row1.addNewTableCell().setText(result.getMvFls().get(0).getLast_name() + "\n" + result.getMvFls().get(0).getFirst_name() + "\n" + result.getMvFls().get(0).getPatronymic());
                    row1.addNewTableCell().setText(result.getMvFls().get(0).isIs_resident() ? "ДА" : "НЕТ");
                    row1.addNewTableCell().setText(result.getMvFls().get(0).getNationality_ru_name());
                    row1.addNewTableCell().setText(result.getMvFls().get(0).getDeath_date());
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Mv_Fl table add exception");
            }
            try {
                if (result.getMvRnOlds() != null || result.getMvRnOlds().size() < 0) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Адреса прописки", Arrays.asList(
                            "Страна",
                            "Город",
                            "Адрес",
                            "Регион",
                            "Дата прописки"
                    ));
                    XWPFTableRow row2 = table.createRow();
                    for (RegAddressFl regAddressFl : result.getRegAddressFls()) {
                        row2.addNewTableCell().setText(regAddressFl.getCountry());
                        row2.addNewTableCell().setText(regAddressFl.getCity());
                        row2.addNewTableCell().setText(regAddressFl.getDistrict());
                        row2.addNewTableCell().setText(regAddressFl.getRegion());
                        row2.addNewTableCell().setText(regAddressFl.getReg_date());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("MV_Rn_Old table add exception");
            }
            try {
                List<MvIinDoc> docs = result.getMvIinDocs();
                if (docs != null && !docs.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Документы", Arrays.asList("Типа Документа", "Орган выдачи", "Дата выдачи", "Срок до", "Номер документа"));
                    for (MvIinDoc doci : docs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(doci.getDoc_type_ru_name());
                        dataRow.addNewTableCell().setText(doci.getIssue_organization_ru_name());
                        dataRow.addNewTableCell().setText(doci.getIssue_date().toString());
                        dataRow.addNewTableCell().setText(doci.getExpiry_date().toString());
                        dataRow.addNewTableCell().setText(doci.getDoc_type_ru_name());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("MV_Iin_Doc table add exception");
            }
            try {
                List<School> schools = result.getSchools();
                if (schools != null && !schools.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Школы", Arrays.asList("БИН", "Название школы", "Класс", "Дата поступления", "Дата окончания"));
                    for (School school : schools) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(school.getSchool_code());
                        dataRow.addNewTableCell().setText(school.getSchool_name());
                        dataRow.addNewTableCell().setText(school.getGrade());
                        dataRow.addNewTableCell().setText(school.getStart_date().toString());
                        dataRow.addNewTableCell().setText(school.getEnd_date().toString());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding school table: " + e.getMessage());
            }

            try {
                List<Universities> universities = result.getUniversities();
                if (universities != null && !universities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Вузы", Arrays.asList("БИН", "Название вуза", "Специализация", "Дата поступления", "Дата окончания", "Длительность обучения", "Курс"));
                    for (Universities university : universities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(university.getStudy_code());
                        dataRow.addNewTableCell().setText(university.getStudy_name());
                        dataRow.addNewTableCell().setText(university.getSpec_name());
                        dataRow.addNewTableCell().setText(university.getStart_date() != null ? university.getStart_date().toString() : "");
                        dataRow.addNewTableCell().setText(university.getEnd_date() != null ? university.getEnd_date().toString() : "");
                        dataRow.addNewTableCell().setText(university.getDuration());
                        dataRow.addNewTableCell().setText(university.getCourse());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding university table: " + e.getMessage());
            }

            try {
                List<MvAutoFl> autos = result.getMvAutoFls();
                if (autos != null && !autos.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Транспорт", Arrays.asList("№", "Статус", "Регистрационный номер", "Марка модель", "Дата выдачи свидетельства", "Дата снятия", "Год выпуска", "Категория", "VIN/Кузов/Шасси", "Серия"));
                    int number = 1;
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
                    for (MvAutoFl auto : autos) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(formatter.format(auto.getEnd_date()).compareTo(formatter.format(new java.util.Date())) > 0 ? "Действителен" : "Не действителен");
                        dataRow.addNewTableCell().setText(auto.getReg_number());
                        dataRow.addNewTableCell().setText(auto.getBrand_model());
                        dataRow.addNewTableCell().setText(auto.getDate_certificate().toString());
                        dataRow.addNewTableCell().setText(auto.getEnd_date().toString());
                        dataRow.addNewTableCell().setText(auto.getRelease_year_tc());
                        dataRow.addNewTableCell().setText(auto.getOwner_category());
                        dataRow.addNewTableCell().setText(auto.getVin_kuzov_shassi());
                        dataRow.addNewTableCell().setText(auto.getSeries_reg_number());
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding auto table: " + e.getMessage());
            }

            try {
                List<FlRelativiesDTO> fl_relatives = result.getFl_relatives();
                if (fl_relatives != null && !fl_relatives.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Родственные связи", Arrays.asList("№", "Статус по отношению к родственнику", "ФИО", "ИИН", "Дата рождения", "Дата регистрации брака", "Дата расторжения брака"));
                    int number = 1;
                    for (FlRelativiesDTO relative : fl_relatives) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(relative.getRelative_type());
                        dataRow.addNewTableCell().setText(relative.getParent_fio());
                        dataRow.addNewTableCell().setText(relative.getParent_iin() != null ? relative.getParent_iin() : "");
                        dataRow.addNewTableCell().setText(relative.getParent_birth_date() != null ? relative.getParent_birth_date().substring(0, 10) : "");
                        dataRow.addNewTableCell().setText(relative.getMarriage_reg_date() != null ? relative.getMarriage_reg_date() : "");
                        dataRow.addNewTableCell().setText(relative.getMarriage_divorce_date() != null ? relative.getMarriage_divorce_date() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding relatives table: " + e.getMessage());
            }
            try {
                List<FlContacts> contacts = result.getContacts();
                if (contacts != null && !contacts.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Контактные данные ФЛ", Arrays.asList("№", "Телефон", "Почта", "Источник"));
                    int number = 1;
                    for (FlContacts contact : contacts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(contact.getPhone());
                        dataRow.addNewTableCell().setText(contact.getEmail());
                        dataRow.addNewTableCell().setText(contact.getSource());
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding contacts table: " + e.getMessage());
            }

            try {
                List<MillitaryAccount> militaryAccounts = result.getMillitaryAccounts();
                if (militaryAccounts != null && !militaryAccounts.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Воинский учет", Arrays.asList("№", "БИН воинской части", "Дата службы с", "Дата службы по"));
                    int number = 1;
                    for (MillitaryAccount account : militaryAccounts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(account.getBin());
                        dataRow.addNewTableCell().setText(account.getDate_start());
                        dataRow.addNewTableCell().setText(account.getDate_end());
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding military accounts table: " + e.getMessage());
            }

            try {
                List<ConvictsJustified> convictsJustifieds = result.getConvictsJustifieds();
                if (convictsJustifieds != null && !convictsJustifieds.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Наименование риска: \"Осужденные\" Количество найденных инф: " + convictsJustifieds.size(),
                            Arrays.asList("№", "Дата рассмотрения в суде 1 инстанции", "Суд 1 инстанции", "Решение по лицу", "Мера наказания по договору", "Квалификация"));
                    int number = 1;
                    for (ConvictsJustified r : convictsJustifieds) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getReg_date() != null ? r.getReg_date() : "");
                        dataRow.addNewTableCell().setText(r.getCourt_of_first_instance() != null ? r.getCourt_of_first_instance() : "");
                        dataRow.addNewTableCell().setText(r.getDecision_on_person() != null ? r.getDecision_on_person() : "");
                        dataRow.addNewTableCell().setText(r.getMeasure_punishment() != null ? r.getMeasure_punishment() : "");
                        dataRow.addNewTableCell().setText(r.getQualification() != null ? r.getQualification() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding convicts justified table: " + e.getMessage());
            }

            try {
                List<ConvictsTerminatedByRehab> convictsTerminatedByRehabs = result.getConvictsTerminatedByRehabs();
                if (convictsTerminatedByRehabs != null && !convictsTerminatedByRehabs.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Административные штрафы",
                            Arrays.asList("№", "Орган выявивший правонарушение", "Дата заведения", "Квалификация", "Решение", "Уровень тяжести"));
                    int number = 1;
                    for (ConvictsTerminatedByRehab r : convictsTerminatedByRehabs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getInvestigative_authority() != null ? r.getInvestigative_authority() : "");
                        dataRow.addNewTableCell().setText(r.getLast_solution_date() != null ? r.getLast_solution_date() : "");
                        dataRow.addNewTableCell().setText(r.getQualification_desc() != null ? r.getQualification_desc() : "");
                        dataRow.addNewTableCell().setText(r.getLast_solution() != null ? r.getLast_solution() : "");
                        dataRow.addNewTableCell().setText(r.getQualification_by_11() != null ? r.getQualification_by_11() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding convicts terminated by rehab table: " + e.getMessage());
            }
            try {
                List<BlockEsf> blockEsfs = result.getBlockEsfs();
                if (blockEsfs != null && !blockEsfs.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Блокировка ЭСФ", Arrays.asList("№", "Дата блокировки", "Дата востановления", "Дата обновления"));
                    int number = 1;
                    for (BlockEsf r : blockEsfs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getStart_dt() != null ? r.getStart_dt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getEnd_dt() != null ? r.getEnd_dt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getUpdate_dt() != null ? r.getUpdate_dt().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding block esf table: " + e.getMessage());
            }

            try {
                List<MvUlFounderFl> mvUlFounderFls = result.getMvUlFounderFls();
                if (mvUlFounderFls != null && !mvUlFounderFls.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения об участниках ЮЛ", Arrays.asList("№", "БИН", "Наименование ЮЛ", "Дата регистрации"));
                    int number = 1;
                    for (MvUlFounderFl r : mvUlFounderFls) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getBin_org() != null ? r.getBin_org() : "");
                        try {
                            dataRow.addNewTableCell().setText(mvUlRepo.getNameByBin(r.getBin_org()));
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }
                        dataRow.addNewTableCell().setText(r.getReg_date() != null ? r.getReg_date().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding mv ul founder fl table: " + e.getMessage());
            }

            try {
                List<NdsEntity> ndsEntities = result.getNdsEntities();
                if (ndsEntities != null && !ndsEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "НДС", Arrays.asList("№", "Дата начала", "Дата конца", "Дата обновления", "Причина"));
                    int number = 1;
                    for (NdsEntity r : ndsEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getStartDt() != null ? r.getStartDt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getEndDt() != null ? r.getEndDt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getUpdateDt() != null ? r.getUpdateDt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getReason() != null ? r.getReason() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding nds entities table: " + e.getMessage());
            }

            try {
                List<IpgoEmailEntity> ipgoEmailEntities = result.getIpgoEmailEntities();
                if (ipgoEmailEntities != null && !ipgoEmailEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по ИПГО", Arrays.asList("№", "Департамент", "Должность", "ИПГО почта"));
                    int number = 1;
                    for (IpgoEmailEntity r : ipgoEmailEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getOrgan() != null ? r.getOrgan().toString() : "");
                        dataRow.addNewTableCell().setText(r.getPosition() != null ? r.getPosition() : "");
                        dataRow.addNewTableCell().setText(r.getEmail() != null ? r.getEmail().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding ipgo email entities table: " + e.getMessage());
            }
            doc.write(baos);
            baos.close();
        }
    }

    public void setCellPadding(XWPFTableCell cell, int top, int left, int bottom, int right) {
        CTTcPr tcPr = cell.getCTTc().addNewTcPr();

        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcMar cellMar = tcPr.isSetTcMar() ? tcPr.getTcMar() : tcPr.addNewTcMar();
        cellMar.addNewTop().setW(BigInteger.valueOf(top));
        cellMar.addNewLeft().setW(BigInteger.valueOf(left));
        cellMar.addNewBottom().setW(BigInteger.valueOf(bottom));
        cellMar.addNewRight().setW(BigInteger.valueOf(right));
    }

    public void generateUl(NodesUL result, ByteArrayOutputStream baos) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            try {
                if (result.getMvUls() != null && !result.getMvUls().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения о юридическом лице", Arrays.asList("№",
                            "БИН",
                            "Наименование организации",
                            "Наименование ОКЭД",
                            "Статус ЮЛ"));
                    int number = 1;
                    for (MvUl a : result.getMvUls()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getBin());
                        dataRow.addNewTableCell().setText(a.getOked());
                        dataRow.addNewTableCell().setText(a.getUl_status());
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding МвUl table: " + e.getMessage());
            }

            try {
                if (result.getMvUlFounderFls() != null && !result.getMvUlFounderFls().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения об участниках ЮЛ", Arrays.asList(
                            "№",
                            "БИН",
                            "Наименование ЮЛ",
                            "Дата регистрации"));

                    int number = 1;
                    for (MvUlFounderFl a : result.getMvUlFounderFls()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        try {
                            String name = mvUlRepo.getNameByBin(a.getBin_org());
                            dataRow.addNewTableCell().setText(name);
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        try {
                            dataRow.addNewTableCell().setText(a.getReg_date().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет даты");
                        }
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding МвUlFounderFl table: " + e.getMessage());
            }
            try {
                if (result.getAccountantListEntities() != null && !result.getAccountantListEntities().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Список бухгалтеров", Arrays.asList(
                            "№",
                            "ИИН",
                            "Проф.",
                            "Фамилия",
                            "Имя"));
                    int number = 1;
                    for (AccountantListEntity a : result.getAccountantListEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getIin());
                        dataRow.addNewTableCell().setText(a.getProf());
                        dataRow.addNewTableCell().setText(a.getLname());
                        dataRow.addNewTableCell().setText(a.getFname());
                        number++;
                    }

                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding AccountantListEntities table: " + e.getMessage());
            }
            doc.write(baos);
            baos.close();
        }
    }
}