package kz.dossier.tools;

import kz.dossier.modelsDossier.*;
import kz.dossier.modelsRisk.*;
import kz.dossier.repositoryDossier.FlPensionMiniRepo;
import kz.dossier.repositoryDossier.MvUlRepo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
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
            CTDocument1 document = doc.getDocument();
            CTBody body = document.getBody();

            if (!body.isSetSectPr()) {
                body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();

            if(!section.isSetPgSz()) {
                section.addNewPgSz();
            }
            CTPageSz pageSize = section.getPgSz();

            pageSize.setW(BigInteger.valueOf(15840));
            pageSize.setH(BigInteger.valueOf(12240));

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

//            try {
//                List<FlRelativiesDTO> fl_relatives = result.getFl_relatives();
//                if (fl_relatives != null && !fl_relatives.isEmpty()) {
//                    XWPFTable table = doc.createTable();
//                    makeTableByProperties(doc, table, "Родственные связи", Arrays.asList("№", "Статус по отношению к родственнику", "ФИО", "ИИН", "Дата рождения", "Дата регистрации брака", "Дата расторжения брака"));
//
//                    int number = 1;
//                    for (FlRelativiesDTO relative : fl_relatives) {
//                        XWPFTableRow dataRow = table.createRow();
//                        dataRow.addNewTableCell().setText(String.valueOf(number));
//                        dataRow.addNewTableCell().setText(relative.getRelative_type());
//                        dataRow.addNewTableCell().setText(relative.getParent_fio());
//                        dataRow.addNewTableCell().setText(relative.getParent_iin() != null ? relative.getParent_iin() : "");
//                        dataRow.addNewTableCell().setText(relative.getParent_birth_date() != null ? relative.getParent_birth_date().substring(0, 10) : "");
//                        dataRow.addNewTableCell().setText(relative.getMarriage_reg_date() != null ? relative.getMarriage_reg_date() : "");
//                        dataRow.addNewTableCell().setText(relative.getMarriage_divorce_date() != null ? relative.getMarriage_divorce_date() : "");
//                        number++;
//                    }
//                    setMarginBetweenTables(doc);
//                }
//            } catch (Exception e) {
//                System.out.println("Exception while adding relatives table: " + e.getMessage());
//            }
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
            try {
                List<Bankrot> bankrotEntities = result.getBankrots();
                if (bankrotEntities != null && !bankrotEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по банкротам", Arrays.asList("№", "ИИН/БИН", "Документ", "Дата обновления", "Причина"));
                    
                    int number = 1;
                    for (Bankrot r : bankrotEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getIin_bin() != null ? r.getIin_bin() : "");
                        dataRow.addNewTableCell().setText(r.getDocument() != null ? r.getDocument() : "");
                        dataRow.addNewTableCell().setText(r.getUpdate_dt() != null ? r.getUpdate_dt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getReason() != null ? r.getReason() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                    // Save the document as needed
                }
            } catch (Exception e) {
                System.out.println("Exception while adding bankrot entities table: " + e.getMessage());
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
                if (result.getFirstCreditBureauEntities() != null && !result.getFirstCreditBureauEntities().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по кредитным бюро", Arrays.asList(
                            "№", "Тип", "Кредит в FOID", "Регион", "Количество FPD SPD", "Сумма долга", "Макс. задержка дней", "Фин. учреждения", "Общее количество кредитов"));

                    
                    int number = 1;
                    for (FirstCreditBureauEntity entity : result.getFirstCreditBureauEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(entity.getType() != null ? entity.getType() : "");
                        dataRow.addNewTableCell().setText(entity.getCreditInFoid() != null ? entity.getCreditInFoid().toString() : "");
                        dataRow.addNewTableCell().setText(entity.getRegion() != null ? entity.getRegion() : "");
                        dataRow.addNewTableCell().setText(entity.getQuantityFpdSpd() != null ? entity.getQuantityFpdSpd().toString() : "");
                        dataRow.addNewTableCell().setText(entity.getAmountOfDebt() != null ? entity.getAmountOfDebt().toString() : "");
                        dataRow.addNewTableCell().setText(entity.getMaxDelayDayNum1() != null ? entity.getMaxDelayDayNum1().toString() : "");
                        dataRow.addNewTableCell().setText(entity.getFinInstitutionsName() != null ? entity.getFinInstitutionsName() : "");
                        dataRow.addNewTableCell().setText(entity.getTotalCountOfCredits() != null ? entity.getTotalCountOfCredits().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                    // Save the document as needed
                }
            } catch (Exception e) {
                System.out.println("Exception while adding first credit bureau entities table: " + e.getMessage());
            }
            try {
                if (result.getAmoral() != null && !result.getAmoral().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по аморальному образу жизни", Arrays.asList("№", "Орган выявивший", "Гражданство", "Дата решения", "Сумма штрафа"));

                    
                    int number = 1;
                    for (ImmoralLifestyle r : result.getAmoral()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getAuthority_detected() != null ? r.getAuthority_detected() : "");
                        dataRow.addNewTableCell().setText(r.getCitizenship_id() != null ? r.getCitizenship_id() : "");
                        dataRow.addNewTableCell().setText(r.getDecision_date() != null ? r.getDecision_date().toString() : "");
                        dataRow.addNewTableCell().setText(r.getFine_amount() != null ? r.getFine_amount().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                    // Save the document as needed
                }
            } catch (Exception e) {
                System.out.println("Exception while adding immoral lifestyle entities table: " + e.getMessage());
            }try {
                if (result.getMzEntities() != null && !result.getMzEntities().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по МЗ", Arrays.asList("№", "Код болезни", "Регистрация", "Статус МЗ", "Медицинская организация"));

                    
                    int number = 1;
                    for (MzEntity r : result.getMzEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getDiseaseCode() != null ? r.getDiseaseCode() : "");
                        dataRow.addNewTableCell().setText(r.getReg() != null ? r.getReg() : "");
                        dataRow.addNewTableCell().setText(r.getStatusMz() != null ? r.getStatusMz() : "");
                        dataRow.addNewTableCell().setText(r.getMedicalOrg() != null ? r.getMedicalOrg() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding MZ entities table: " + e.getMessage());
            }try {
                if (result.getWantedListEntities() != null && !result.getWantedListEntities().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по разыскиваемым", Arrays.asList("№", "Дни", "Орган", "Статус", "Дата актуальности"));

                    
                    int number = 1;
                    for (WantedListEntity r : result.getWantedListEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getDays() != null ? r.getDays().toString() : "");
                        dataRow.addNewTableCell().setText(r.getOrgan() != null ? r.getOrgan() : "");
                        dataRow.addNewTableCell().setText(r.getStatus() != null ? r.getStatus() : "");
                        dataRow.addNewTableCell().setText(r.getRelevanceDate() != null ? r.getRelevanceDate().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                    // Save the document as needed
                }
            } catch (Exception e) {
                System.out.println("Exception while adding wanted list entities table: " + e.getMessage());
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
            CTDocument1 document = doc.getDocument();
            CTBody body = document.getBody();

            if (!body.isSetSectPr()) {
                body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();

            if(!section.isSetPgSz()) {
                section.addNewPgSz();
            }
            CTPageSz pageSize = section.getPgSz();

            pageSize.setW(BigInteger.valueOf(15840));
            pageSize.setH(BigInteger.valueOf(12240));
            try {
                if (result.getMvUls() != null && !result.getMvUls().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения о юридическом лице", Arrays.asList(
                            "БИН",
                            "Наименование организации",
                            "Наименование ОКЭД",
                            "Статус ЮЛ"
                    ));

                    for (MvUl a : result.getMvUls()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(a.getBin() != null ? a.getBin() : "");
                        dataRow.addNewTableCell().setText(a.getFull_name_kaz() != null ? a.getFull_name_kaz() : "");
                        dataRow.addNewTableCell().setText(a.getHead_organization() != null ? a.getHead_organization() : "");
                        dataRow.addNewTableCell().setText(a.getUl_status() != null ? a.getUl_status() : "");
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding MV UL table: " + e.getMessage());
            }try {
                if (result.getMvUlFounderFls() != null && !result.getMvUlFounderFls().isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения об участниках ЮЛ", Arrays.asList(
                            "№",
                            "БИН",
                            "Наименование ЮЛ",
                            "Дата регистрации"
                    ));

                    int number = 1;
                    for (MvUlFounderFl a : result.getMvUlFounderFls()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));

                        try {
                            String name = mvUlRepo.getNameByBin(a.getBin_org());
                            dataRow.addNewTableCell().setText(name != null ? name : "Нет");
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }

                        try {
                            dataRow.addNewTableCell().setText(a.getReg_date() != null ? a.getReg_date().toString() : "Нет даты");
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
            }try {
                List<Omn> omns = result.getOmns();
                if (omns != null && !omns.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "ОМНС", Arrays.asList(
                            "РНН",
                            "Название налогоплательщика",
                            "ФИО налогоплательщика",
                            "ФИО руководителя",
                            "ИИН руководителя",
                            "РНН руководителя"
                    ));

                    int number = 1;
                    for (Omn a : omns) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getRnn() != null ? a.getRnn() : "");
                        dataRow.addNewTableCell().setText(a.getTaxpayer_name() != null ? a.getTaxpayer_name() : "");
                        dataRow.addNewTableCell().setText(a.getTaxpayer_fio() != null ? a.getTaxpayer_fio() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_fio() != null ? a.getLeader_fio() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_iin() != null ? a.getLeader_iin() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_rnn() != null ? a.getLeader_rnn() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding OMNS table: " + e.getMessage());
            }

            try {
                List<Equipment> equipmentList = result.getEquipment();
                if (equipmentList != null && !equipmentList.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Транспорт", Arrays.asList(
                            "Адрес",
                            "Гос. Номер",
                            "Номер серии рег.",
                            "Дата регистрации",
                            "Причина",
                            "VIN",
                            "Спец.",
                            "Тип",
                            "Форма",
                            "Брэнд",
                            "Модель"
                    ));

                    int number = 1;
                    for (Equipment a : equipmentList) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getOwner_address() != null ? a.getOwner_address() : "");
                        dataRow.addNewTableCell().setText(a.getGov_number() != null ? a.getGov_number() : "");
                        dataRow.addNewTableCell().setText(a.getReg_series_num() != null ? a.getReg_series_num() : "");
                        dataRow.addNewTableCell().setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.addNewTableCell().setText(a.getReg_reason() != null ? a.getReg_reason() : "");
                        dataRow.addNewTableCell().setText(a.getVin() != null ? a.getVin() : "");
                        dataRow.addNewTableCell().setText(a.getEquipment_spec() != null ? a.getEquipment_spec() : "");
                        dataRow.addNewTableCell().setText(a.getEquipment_type() != null ? a.getEquipment_type() : "");
                        dataRow.addNewTableCell().setText(a.getEquipment_form() != null ? a.getEquipment_form() : "");
                        dataRow.addNewTableCell().setText(a.getBrand() != null ? a.getBrand() : "");
                        dataRow.addNewTableCell().setText(a.getEquipment_model() != null ? a.getEquipment_model() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Equipment table: " + e.getMessage());
            }

            try {
                List<Msh> mshes = result.getMshes();
                if (mshes != null && !mshes.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "МШЭС", Arrays.asList(
                            "Тип оборудования",
                            "Модель оборудования",
                            "VIN",
                            "Гос. номер",
                            "Дата регистрации"
                    ));

                    int number = 1;
                    for (Msh a : mshes) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getEquipmentType() != null ? a.getEquipmentType() : "");
                        dataRow.addNewTableCell().setText(a.getEquipmentModel() != null ? a.getEquipmentModel() : "");
                        dataRow.addNewTableCell().setText(a.getVin() != null ? a.getVin() : "");
                        dataRow.addNewTableCell().setText(a.getGovNumber() != null ? a.getGovNumber() : "");
                        try {
                            dataRow.addNewTableCell().setText(a.getRegDate() != null ? a.getRegDate().toString() : "Нет");
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding MSHES table: " + e.getMessage());
            }
            try {
                List<Dormant> dormants = result.getDormants();
                if (dormants != null && !dormants.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Дорманс", Arrays.asList(
                            "№",
                            "РНН",
                            "Название налогоплательщика",
                            "ФИО налогоплательщика",
                            "ФИО руководителя",
                            "ИИН руководителя",
                            "РНН руководителя",
                            "Дата заказа"
                    ));

                    int number = 1;
                    for (Dormant a : dormants) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getRnn() != null ? a.getRnn() : "");
                        dataRow.addNewTableCell().setText(a.getTaxpayer_name() != null ? a.getTaxpayer_name() : "");
                        dataRow.addNewTableCell().setText(a.getTaxpayer_fio() != null ? a.getTaxpayer_fio() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_fio() != null ? a.getLeader_fio() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_iin() != null ? a.getLeader_iin() : "");
                        dataRow.addNewTableCell().setText(a.getLeader_rnn() != null ? a.getLeader_rnn() : "");
                        dataRow.addNewTableCell().setText(a.getOrder_date() != null ? a.getOrder_date() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Dormant table: " + e.getMessage());
            }

            try {
                List<Bankrot> bankrots = result.getBankrots();
                if (bankrots != null && !bankrots.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Банкроты", Arrays.asList(
                            "№",
                            "Документ",
                            "Дата обновления",
                            "Причина"
                    ));

                    int number = 1;
                    for (Bankrot a : bankrots) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getDocument() != null ? a.getDocument() : "");
                        try {
                            dataRow.addNewTableCell().setText(a.getUpdate_dt() != null ? a.getUpdate_dt().toString() : "Дата отсутствует");
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Дата отсутствует");
                        }
                        dataRow.addNewTableCell().setText(a.getReason() != null ? a.getReason() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Bankrot table: " + e.getMessage());
            }

            try {
                List<Adm> adms = result.getAdms();
                if (adms != null && !adms.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Администрация", Arrays.asList(
                            "№",
                            "Номер материала",
                            "Дата регистрации",
                            "15",
                            "16",
                            "17",
                            "Наименование юр. лица",
                            "Адрес юр. лица",
                            "Марка автомобиля",
                            "Гос. Номер авто"
                    ));

                    int number = 1;
                    for (Adm a : adms) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getMaterial_num() != null ? a.getMaterial_num() : "");
                        dataRow.addNewTableCell().setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.addNewTableCell().setText(a.getFifteen() != null ? a.getFifteen() : "");
                        dataRow.addNewTableCell().setText(a.getSixteen() != null ? a.getSixteen() : "");
                        dataRow.addNewTableCell().setText(a.getSeventeen() != null ? a.getSeventeen() : "");
                        dataRow.addNewTableCell().setText(a.getUl_org_name() != null ? a.getUl_org_name() : "");
                        dataRow.addNewTableCell().setText(a.getUl_adress() != null ? a.getUl_adress() : "");
                        dataRow.addNewTableCell().setText(a.getVehicle_brand() != null ? a.getVehicle_brand() : "");
                        dataRow.addNewTableCell().setText(a.getState_auto_num() != null ? a.getState_auto_num() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Adm table: " + e.getMessage());
            }

            try {
                List<Criminals> criminals = result.getCriminals();
                if (criminals != null && !criminals.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Преступления", Arrays.asList(
                            "№",
                            "Наименование суда",
                            "Дата судебного решения",
                            "Решение",
                            "Название преступления",
                            "Приговор",
                            "Дополнительная информация",
                            "Обращение",
                            "ЕРДР"
                    ));

                    int number = 1;
                    for (Criminals a : criminals) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getCourt_name() != null ? a.getCourt_name() : "");
                        dataRow.addNewTableCell().setText(a.getCourt_dt() != null ? a.getCourt_dt() : "");
                        dataRow.addNewTableCell().setText(a.getDecision() != null ? a.getDecision() : "");
                        dataRow.addNewTableCell().setText(a.getCrime_name() != null ? a.getCrime_name() : "");
                        dataRow.addNewTableCell().setText(a.getSentence() != null ? a.getSentence() : "");
                        dataRow.addNewTableCell().setText(a.getAdd_info() != null ? a.getAdd_info() : "");
                        dataRow.addNewTableCell().setText(a.getTreatment() != null ? a.getTreatment() : "");
                        dataRow.addNewTableCell().setText(a.getErdr() != null ? a.getErdr() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Criminals table: " + e.getMessage());
            }try {
                List<BlockEsf> blockEsfs = result.getBlockEsfs();
                if (blockEsfs != null && !blockEsfs.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Блокировка ЕСФ", Arrays.asList(
                            "№",
                            "Дата начала",
                            "Дата окончания",
                            "Дата обновления"
                    ));

                    int number = 1;
                    for (BlockEsf a : blockEsfs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        try {
                            dataRow.addNewTableCell().setText(a.getStart_dt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        try {
                            dataRow.addNewTableCell().setText(a.getEnd_dt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        try {
                            dataRow.addNewTableCell().setText(a.getUpdate_dt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding BlockEsf table: " + e.getMessage());
            }

            try {
                List<NdsEntity> ndsEntities = result.getNdsEntities();
                if (ndsEntities != null && !ndsEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Объекты НДС", Arrays.asList(
                            "№",
                            "Дата начала",
                            "Дата окончания",
                            "Причина",
                            "Дата обновления"
                    ));

                    int number = 1;
                    for (NdsEntity a : ndsEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        try {
                            dataRow.addNewTableCell().setText(a.getStartDt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        try {
                            dataRow.addNewTableCell().setText(a.getEndDt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        dataRow.addNewTableCell().setText(a.getReason() != null ? a.getReason() : "");
                        try {
                            dataRow.addNewTableCell().setText(a.getUpdateDt().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("Нет");
                        }
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding NdsEntity table: " + e.getMessage());
            }

            try {
                List<MvRnOld> mvRnOlds = result.getMvRnOlds();
                if (mvRnOlds != null && !mvRnOlds.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "mv_rn_old", Arrays.asList(
                            "№",
                            "Назначение использования",
                            "Статус недвижимости",
                            "Адрес",
                            "История адресов",
                            "Тип собственности",
                            "Вид собственности",
                            "Статус характеристики недвижимости",
                            "Дата регистрации в реестре",
                            "Дата окончания регистрации",
                            "Возникновение права в реестре",
                            "Статус в реестре"
                    ));

                    int number = 1;
                    for (MvRnOld a : mvRnOlds) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getIntended_use_rus() != null ? a.getIntended_use_rus() : "");
                        dataRow.addNewTableCell().setText(a.getEstate_status_rus() != null ? a.getEstate_status_rus() : "");
                        dataRow.addNewTableCell().setText(a.getAddress_rus() != null ? a.getAddress_rus() : "");
                        dataRow.addNewTableCell().setText(a.getAddress_history_rus() != null ? a.getAddress_history_rus() : "");
                        dataRow.addNewTableCell().setText(a.getType_of_property_rus() != null ? a.getType_of_property_rus() : "");
                        dataRow.addNewTableCell().setText(a.getProperty_type_rus() != null ? a.getProperty_type_rus() : "");
                        dataRow.addNewTableCell().setText(a.getEstate_characteristic_status_rus() != null ? a.getEstate_characteristic_status_rus() : "");
                        dataRow.addNewTableCell().setText(a.getRegister_reg_date() != null ? a.getRegister_reg_date() : "");
                        dataRow.addNewTableCell().setText(a.getRegister_end_date() != null ? a.getRegister_end_date() : "");
                        dataRow.addNewTableCell().setText(a.getRegister_emergence_rights_rus() != null ? a.getRegister_emergence_rights_rus() : "");
                        dataRow.addNewTableCell().setText(a.getRegister_status_rus() != null ? a.getRegister_status_rus() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding MvRnOld table: " + e.getMessage());
            }

            try {
                List<FpgTempEntity> fpgTempEntities = result.getFpgTempEntities();
                if (fpgTempEntities != null && !fpgTempEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Временные объекты ФПГ", Arrays.asList(
                            "№",
                            "Бенефициар"
                    ));

                    int number = 1;
                    for (FpgTempEntity a : fpgTempEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getBeneficiary() != null ? a.getBeneficiary() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding FpgTempEntity table: " + e.getMessage());
            }try {
                List<Pdl> pdls = result.getPdls();
                if (pdls != null && !pdls.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "ПДЛ", Arrays.asList(
                            "№",
                            "ИИН",
                            "Полное наименование организации",
                            "ФИО",
                            "Орган",
                            "Область",
                            "ФИО супруг(и)",
                            "Орган супруг(и)",
                            "Должность супруга",
                            "ИИН супруга"
                    ));
                    int number = 1;
                    for (Pdl a : pdls) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getIin() != null ? a.getIin() : "");
                        dataRow.addNewTableCell().setText(a.getOrganization_fullname() != null ? a.getOrganization_fullname() : "");
                        dataRow.addNewTableCell().setText(a.getFio() != null ? a.getFio() : "");
                        dataRow.addNewTableCell().setText(a.getOrgan() != null ? a.getOrgan() : "");
                        dataRow.addNewTableCell().setText(a.getOblast() != null ? a.getOblast() : "");
                        dataRow.addNewTableCell().setText(a.getSpouse_fio() != null ? a.getSpouse_fio() : "");
                        dataRow.addNewTableCell().setText(a.getSpouse_organ() != null ? a.getSpouse_organ() : "");
                        dataRow.addNewTableCell().setText(a.getSpouse_position() != null ? a.getSpouse_position() : "");
                        dataRow.addNewTableCell().setText(a.getSpouse_iin() != null ? a.getSpouse_iin() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Pdl table: " + e.getMessage());
            }

            try {
                List<CommodityProducer> commodityProducers = result.getCommodityProducers();
                if (commodityProducers != null && !commodityProducers.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Производители товаров", Arrays.asList(
                            "№",
                            "Наименование ССП",
                            "Количество",
                            "Производитель",
                            "Статус",
                            "Регион",
                            "СЗТП"
                    ));
                    int number = 1;
                    for (CommodityProducer a : commodityProducers) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getSspName() != null ? a.getSspName() : "");
                        dataRow.addNewTableCell().setText(String.valueOf(a.getCount()));
                        dataRow.addNewTableCell().setText(a.getProducer() != null ? a.getProducer() : "");
                        dataRow.addNewTableCell().setText(a.getStatus() != null ? a.getStatus() : "");
                        dataRow.addNewTableCell().setText(a.getRegion() != null ? a.getRegion() : "");
                        dataRow.addNewTableCell().setText(a.getSztp() != null ? a.getSztp() : "");
                    number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding CommodityProducer table: " + e.getMessage());
            }

            try {
                RegAddressUlEntity regAddressUlEntity = result.getRegAddressUlEntities();
                if (regAddressUlEntity != null) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Адрес", Arrays.asList(
                            "Дата регистрации",
                            "Название организации (на русском)",
                            "Регион регистрации (на русском)",
                            "Район регистрации (на русском)",
                            "Сельский район регистрации (на русском)",
                            "Населенный пункт регистрации (на русском)",
                            "Улица регистрации (на русском)",
                            "Номер здания",
                            "Номер блока",
                            "Номер корпуса здания",
                            "Офис (номер)",
                            "Название ОКЭД (на русском)",
                            "Статус ЮЛ",
                            "Активный"
                    ));

                    XWPFTableRow dataRow = table.createRow();
                    try {
                        dataRow.addNewTableCell().setText(regAddressUlEntity.getRegDate().toString());
                    } catch (Exception e) {
                        dataRow.addNewTableCell().setText("");
                    }
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getOrgNameRu() != null ? regAddressUlEntity.getOrgNameRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrRegionRu() != null ? regAddressUlEntity.getRegAddrRegionRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrDistrictRu() != null ? regAddressUlEntity.getRegAddrDistrictRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrRuralDistrictRu() != null ? regAddressUlEntity.getRegAddrRuralDistrictRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrLocalityRu() != null ? regAddressUlEntity.getRegAddrLocalityRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrStreetRu() != null ? regAddressUlEntity.getRegAddrStreetRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrBuildingNum() != null ? regAddressUlEntity.getRegAddrBuildingNum() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrBlockNum() != null ? regAddressUlEntity.getRegAddrBlockNum() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrBuildingBodyNum() != null ? regAddressUlEntity.getRegAddrBuildingBodyNum() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getRegAddrOffice() != null ? regAddressUlEntity.getRegAddrOffice() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getOkedNameRu() != null ? regAddressUlEntity.getOkedNameRu() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getUl_status() != null ? regAddressUlEntity.getUl_status() : "");
                    dataRow.addNewTableCell().setText(regAddressUlEntity.getActive() ? "Активен" : "Неактивен");

                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding RegAddressUlEntity table: " + e.getMessage());
            }

            try {
                List<SvedenyaObUchastnikovUlEntity> svedenyaObUchastnikovUlEntities = result.getSvedenyaObUchastnikovUlEntities();
                if (svedenyaObUchastnikovUlEntities != null && !svedenyaObUchastnikovUlEntities.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения об участниках ЮЛ", Arrays.asList(
                            "№",
                            "ФИО или наименование ЮЛ",
                            "Идентификатор",
                            "Дата регистрации",
                            "Риск"
                    ));
                    int number = 1;
                    for (SvedenyaObUchastnikovUlEntity a : svedenyaObUchastnikovUlEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(a.getFIOorUlName() != null ? a.getFIOorUlName() : "");
                        dataRow.addNewTableCell().setText(a.getIdentificator() != null ? a.getIdentificator() : "");
                        dataRow.addNewTableCell().setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.addNewTableCell().setText(a.getRisk() != null ? a.getRisk() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding SvedenyaObUchastnikovUlEntity table: " + e.getMessage());
            }
            doc.write(baos);
            baos.close();
        }
    }
    private void setTableKeepTogether(XWPFTable table) {
        table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
    }

    private void setRowKeepWithNext(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            cell.getCTTc().getPList().forEach(ctP -> {
                CTPPr ppr = ctP.isSetPPr() ? ctP.getPPr() : ctP.addNewPPr();
                ppr.addNewKeepNext().setVal(STOnOff1.ON);
                ppr.addNewKeepLines().setVal(STOnOff1.ON);
            });
        }
    }
}