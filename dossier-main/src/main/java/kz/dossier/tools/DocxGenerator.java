package kz.dossier.tools;

import kz.dossier.modelsDossier.*;
import kz.dossier.modelsRisk.*;
import kz.dossier.repositoryDossier.FlPensionMiniRepo;
import kz.dossier.repositoryDossier.MvUlRepo;
import kz.dossier.service.MyService;
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
    @Autowired
    MyService myService;
    private MvUlRepo mvUlRepo;

    private XWPFDocument makeTableByProperties(XWPFDocument doc, XWPFTable table, String title, List<String> properties) {
        table.setWidth("100%");
        XWPFTableRow row = table.getRow(0);
        // Create cells for the header row
        for (int i = 0; i < properties.size(); i++) {
            if (i == 0) {
                row.getCell(0).setText(properties.get(i));
            } else {
                row.addNewTableCell().setText(properties.get(i));
            }
        }
        return doc;
    }

    private void setMarginBetweenTables(XWPFDocument doc) {
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run2 = paragraph.createRun();
        run2.addBreak();  // Добавляем перенос строки
        run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
    }

    private void creteTitle(XWPFDocument doc,String title){
        XWPFParagraph titleParagraph = doc.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText(title);
        titleRun.setBold(true);
        titleRun.setFontSize(14);
    }

    public void generateDoccc(NodesFL result, ByteArrayOutputStream baos) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            CTDocument1 document = doc.getDocument();
            CTBody body = document.getBody();

            if (!body.isSetSectPr()) {
                body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();

            if (!section.isSetPgSz()) {
                section.addNewPgSz();
            }
            CTPageSz pageSize = section.getPgSz();

            pageSize.setW(BigInteger.valueOf(15840));
            pageSize.setH(BigInteger.valueOf(12240));

            creteTitle(doc, "Сведения о физическом лице");

            XWPFTable table = doc.createTable();


            XWPFTableRow row1 = table.getRow(0);
            row1.getCell(0).setText("First Row, First Column");
            row1.addNewTableCell().setText("First Row, Second Column");
            row1.addNewTableCell().setText("First Row, Third Column");

            //Creating second Row
            XWPFTableRow row2 = table.createRow();
            row2.getCell(0).setText("Second Row, First Column");
            row2.getCell(1).setText("Second Row, Second Column");
            row2.getCell(2).setText("Second Row, Third Column");

            //create third row
            XWPFTableRow row3 = table.createRow();
            row3.getCell(0).setText("Third Row, First Column");
            row3.getCell(1).setText("Third Row, Second Column");
            row3.getCell(2).setText("Third Row, Third Column");





            doc.write(baos);
            baos.close();

        }
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
                    creteTitle(doc,"Сведения о физическом лице");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Сведения о физическом лице", Arrays.asList(
                            "Фото",
                            "ИИН",
                            "ФИО",
                            "Резидент",
                            "Национальность",
                            "Дата смерти"));
                    
                    XWPFTableRow row1 = table.createRow();
                    XWPFTableCell cell1 = row1.getCell(0);

                    XWPFParagraph paragraph2 = cell1.addParagraph();

                    setCellPadding(cell1, 200, 200, 200, 200);
                    XWPFRun run1 = paragraph2.createRun();

                    byte[] imageBytes = result.getPhotoDbf().get(0).getPhoto();
                    ByteArrayInputStream imageStream = new ByteArrayInputStream(imageBytes);
                    int imageType = XWPFDocument.PICTURE_TYPE_PNG; // Change according to your image type (e.g., PICTURE_TYPE_JPEG)
                    run1.addPicture(imageStream, imageType, "image.png", Units.toEMU(75), Units.toEMU(100));
                    row1.getCell(1).setText(result.getMvFls().get(0).getIin());
                    row1.getCell(2).setText(result.getMvFls().get(0).getLast_name() + "\n" + result.getMvFls().get(0).getFirst_name() + "\n" + result.getMvFls().get(0).getPatronymic());
                    row1.getCell(3).setText(result.getMvFls().get(0).isIs_resident() ? "ДА" : "НЕТ");
                    row1.getCell(4).setText(result.getMvFls().get(0).getNationality_ru_name());
                    row1.getCell(5).setText(result.getMvFls().get(0).getDeath_date() !=null ? result.getMvFls().get(0).getDeath_date() : "Отсутсвует");
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Mv_Fl table add exception");
            }
            try {
                if (result.getMvRnOlds() != null && result.getMvRnOlds().size() > 0) {
                    creteTitle(doc, "Адреса прописки");

                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Адреса прописки", Arrays.asList(
                            "Страна",
                            "Город",
                            "Адрес",
                            "Регион",
                            "Дата прописки"
                    ));

                    // Populate the table with data
                    for (RegAddressFl regAddressFl : result.getRegAddressFls()) {
                        XWPFTableRow row = table.createRow();
                        row.getCell(0).setText(regAddressFl.getCountry());
                        row.getCell(1).setText(regAddressFl.getCity());
                        row.getCell(2).setText(regAddressFl.getDistrict());
                        row.getCell(3).setText(regAddressFl.getRegion());
                        row.getCell(4).setText(regAddressFl.getReg_date());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("MV_Rn_Old table add exception");
            }
            try {
                List<MvIinDoc> docs = result.getMvIinDocs();
                if (docs != null && !docs.isEmpty()) {
                    creteTitle(doc,"Документы");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Документы", Arrays.asList("Типа Документа", "Орган выдачи", "Дата выдачи", "Срок до", "Номер документа"));
                    
                    for (MvIinDoc doci : docs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(doci.getDoc_type_ru_name());
                        dataRow.getCell(1).setText(doci.getIssue_organization_ru_name());
                        dataRow.getCell(2).setText(doci.getIssue_date().toString());
                        dataRow.getCell(3).setText(doci.getExpiry_date().toString());
                        dataRow.getCell(4).setText(doci.getDoc_type_ru_name());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("MV_Iin_Doc table add exception");
            }
            try {
                List<FlRelativiesDTO> flRelativiesDTOS = myService.getRelativesInfo(result.getMvIinDocs().get(0).getIin());
                if (flRelativiesDTOS != null && !flRelativiesDTOS.isEmpty()) {
                    System.out.println(flRelativiesDTOS.size());
                    creteTitle(doc,"Родсвтенники");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Родсвтенники", Arrays.asList("Статус по отношению к родственнику", "ФИО", "Дата регистрации брака", "Дата расторжения брака", "ИИН"));

                    for (FlRelativiesDTO flRelativiesDTO : flRelativiesDTOS) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(flRelativiesDTO.getRelative_type() != null ? flRelativiesDTO.getRelative_type() : "");
                        dataRow.getCell(1).setText(flRelativiesDTO.getParent_fio() != null ? flRelativiesDTO.getParent_fio() : "");
                        dataRow.getCell(2).setText(flRelativiesDTO.getMarriage_reg_date() != null ? flRelativiesDTO.getMarriage_reg_date().toString() : "");
                        dataRow.getCell(3).setText(flRelativiesDTO.getMarriage_divorce_date() != null ? flRelativiesDTO.getMarriage_divorce_date().toString() : "");
                        dataRow.getCell(4).setText(flRelativiesDTO.getParent_iin() != null ? flRelativiesDTO.getParent_iin() : "");
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("flRelativiesDTO table add exception");
            }
            try {
                List<School> schools = result.getSchools();
                if (schools != null && !schools.isEmpty()) {
                    creteTitle(doc,"Школы");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Школы", Arrays.asList("БИН", "Название школы", "Класс", "Дата поступления", "Дата окончания"));
                    
                    for (School school : schools) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(school.getSchool_code());
                        dataRow.getCell(1).setText(school.getSchool_name());
                        dataRow.getCell(2).setText(school.getGrade());
                        dataRow.getCell(3).setText(school.getStart_date().toString());
                        dataRow.getCell(4).setText(school.getEnd_date().toString());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding school table: " + e.getMessage());
            }

            try {
                List<Universities> universities = result.getUniversities();
                if (universities != null && !universities.isEmpty()) {
                    creteTitle(doc,"Вузы");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Вузы", Arrays.asList("БИН", "Название вуза", "Специализация", "Дата поступления", "Дата окончания", "Длительность обучения", "Курс"));
                    
                    for (Universities university : universities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(university.getStudy_code());
                        dataRow.getCell(1).setText(university.getStudy_name());
                        dataRow.getCell(2).setText(university.getSpec_name());
                        dataRow.getCell(3).setText(university.getStart_date() != null ? university.getStart_date().toString() : "");
                        dataRow.getCell(4).setText(university.getEnd_date() != null ? university.getEnd_date().toString() : "");
                        dataRow.getCell(5).setText(university.getDuration());
                        dataRow.getCell(6).setText(university.getCourse());
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding university table: " + e.getMessage());
            }

            try {
                List<MvAutoFl> autos = result.getMvAutoFls();
                if (autos != null && !autos.isEmpty()) {
                    creteTitle(doc,"Транспорт");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Транспорт", Arrays.asList("№", "Статус", "Регистрационный номер", "Марка модель", "Дата выдачи свидетельства", "Дата снятия", "Год выпуска", "Категория", "VIN/Кузов/Шасси", "Серия"));
                    
                    int number = 1;
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
                    for (MvAutoFl auto : autos) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(formatter.format(auto.getEnd_date()).compareTo(formatter.format(new java.util.Date())) > 0 ? "Действителен" : "Не действителен");
                        dataRow.getCell(2).setText(auto.getReg_number());
                        dataRow.getCell(3).setText(auto.getBrand_model());
                        dataRow.getCell(4).setText(auto.getDate_certificate().toString());
                        dataRow.getCell(5).setText(auto.getEnd_date().toString());
                        dataRow.getCell(6).setText(auto.getRelease_year_tc());
                        dataRow.getCell(7).setText(auto.getOwner_category());
                        dataRow.getCell(8).setText(auto.getVin_kuzov_shassi());
                        dataRow.getCell(9).setText(auto.getSeries_reg_number());
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
                    creteTitle(doc,"Контактные данные ФЛ");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Контактные данные ФЛ", Arrays.asList("№", "Телефон", "Почта", "Источник"));
                    
                    int number = 1;
                    for (FlContacts contact : contacts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(contact.getPhone());
                        dataRow.getCell(2).setText(contact.getEmail());
                        dataRow.getCell(3).setText(contact.getSource());
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
                    creteTitle(doc,"Воинский учет");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Воинский учет", Arrays.asList("№", "БИН воинской части", "Дата службы с", "Дата службы по"));
                    
                    int number = 1;
                    for (MillitaryAccount account : militaryAccounts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(account.getBin());
                        dataRow.getCell(2).setText(account.getDate_start());
                        dataRow.getCell(3).setText(account.getDate_end());
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
                    creteTitle(doc,"Сведения об участниках ЮЛ");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения об участниках ЮЛ", Arrays.asList("№", "БИН", "Наименование ЮЛ", "Дата регистрации"));
                    
                    int number = 1;
                    for (MvUlFounderFl r : mvUlFounderFls) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getBin_org() != null ? r.getBin_org() : "");
                        dataRow.getCell(2).setText(mvUlRepo.getNameByBin(r.getBin_org()) != null ? mvUlRepo.getNameByBin(r.getBin_org()) : "");
                        dataRow.getCell(3).setText(r.getReg_date() != null ? r.getReg_date().toString() : "");
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
                    creteTitle(doc,"НДС");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "НДС", Arrays.asList("№", "Дата начала", "Дата конца", "Дата обновления", "Причина"));
                    
                    int number = 1;
                    for (NdsEntity r : ndsEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getStartDt() != null ? r.getStartDt().toString() : "");
                        dataRow.getCell(2).setText(r.getEndDt() != null ? r.getEndDt().toString() : "");
                        dataRow.getCell(3).setText(r.getUpdateDt() != null ? r.getUpdateDt().toString() : "");
                        dataRow.getCell(4).setText(r.getReason() != null ? r.getReason() : "");
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
                    creteTitle(doc,"Сведения по ИПГО");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по ИПГО", Arrays.asList("№", "Департамент", "Должность", "ИПГО почта"));
                    
                    int number = 1;
                    for (IpgoEmailEntity r : ipgoEmailEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getOrgan() != null ? r.getOrgan().toString() : "");
                        dataRow.getCell(2).setText(r.getPosition() != null ? r.getPosition() : "");
                        dataRow.getCell(3).setText(r.getEmail() != null ? r.getEmail().toString() : "");
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
                    creteTitle(doc,"Сведения по банкротам");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по банкротам", Arrays.asList("№", "ИИН/БИН", "Документ", "Дата обновления", "Причина"));
                    
                    int number = 1;
                    for (Bankrot r : bankrotEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getIin_bin() != null ? r.getIin_bin() : "");
                        dataRow.getCell(2).setText(r.getDocument() != null ? r.getDocument() : "");
                        dataRow.getCell(3).setText(r.getUpdate_dt() != null ? r.getUpdate_dt().toString() : "");
                        dataRow.getCell(4).setText(r.getReason() != null ? r.getReason() : "");
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
                    creteTitle(doc,"Наименование риска: \"Осужденные\" Количество найденных инф: " + convictsJustifieds.size());
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Наименование риска: \"Осужденные\" Количество найденных инф: " + convictsJustifieds.size(),
                            Arrays.asList("№", "Дата рассмотрения в суде 1 инстанции", "Суд 1 инстанции", "Решение по лицу", "Мера наказания по договору", "Квалификация"));
                    
                    int number = 1;
                    for (ConvictsJustified r : convictsJustifieds) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getReg_date() != null ? r.getReg_date() : "");
                        dataRow.getCell(2).setText(r.getCourt_of_first_instance() != null ? r.getCourt_of_first_instance() : "");
                        dataRow.getCell(3).setText(r.getDecision_on_person() != null ? r.getDecision_on_person() : "");
                        dataRow.getCell(4).setText(r.getMeasure_punishment() != null ? r.getMeasure_punishment() : "");
                        dataRow.getCell(5).setText(r.getQualification() != null ? r.getQualification() : "");
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
                    creteTitle(doc,"Административные штрафы");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Административные штрафы",
                            Arrays.asList("№", "Орган выявивший правонарушение", "Дата заведения", "Квалификация", "Решение", "Уровень тяжести"));
                    
                    int number = 1;
                    for (ConvictsTerminatedByRehab r : convictsTerminatedByRehabs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getInvestigative_authority() != null ? r.getInvestigative_authority() : "");
                        dataRow.getCell(2).setText(r.getLast_solution_date() != null ? r.getLast_solution_date() : "");
                        dataRow.getCell(3).setText(r.getQualification_desc() != null ? r.getQualification_desc() : "");
                        dataRow.getCell(4).setText(r.getLast_solution() != null ? r.getLast_solution() : "");
                        dataRow.getCell(5).setText(r.getQualification_by_11() != null ? r.getQualification_by_11() : "");
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
                    creteTitle(doc,"Блокировка ЭСФ");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Блокировка ЭСФ", Arrays.asList("№", "Дата блокировки", "Дата востановления", "Дата обновления"));
                    
                    int number = 1;
                    for (BlockEsf r : blockEsfs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getStart_dt() != null ? r.getStart_dt().toString() : "");
                        dataRow.getCell(2).setText(r.getEnd_dt() != null ? r.getEnd_dt().toString() : "");
                        dataRow.getCell(3).setText(r.getUpdate_dt() != null ? r.getUpdate_dt().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding block esf table: " + e.getMessage());
            }
            try {
                if (result.getFirstCreditBureauEntities() != null && !result.getFirstCreditBureauEntities().isEmpty()) {
                    creteTitle(doc,"Сведения по кредитным бюро");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по кредитным бюро", Arrays.asList(
                            "№", "Тип", "Кредит в FOID", "Регион", "Количество FPD SPD", "Сумма долга", "Макс. задержка дней", "Фин. учреждения", "Общее количество кредитов"));

                    
                    int number = 1;
                    for (FirstCreditBureauEntity entity : result.getFirstCreditBureauEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(entity.getType() != null ? entity.getType() : "");
                        dataRow.getCell(2).setText(entity.getCreditInFoid() != null ? entity.getCreditInFoid().toString() : "");
                        dataRow.getCell(3).setText(entity.getRegion() != null ? entity.getRegion() : "");
                        dataRow.getCell(4).setText(entity.getQuantityFpdSpd() != null ? entity.getQuantityFpdSpd().toString() : "");
                        dataRow.getCell(5).setText(entity.getAmountOfDebt() != null ? entity.getAmountOfDebt().toString() : "");
                        dataRow.getCell(6).setText(entity.getMaxDelayDayNum1() != null ? entity.getMaxDelayDayNum1().toString() : "");
                        dataRow.getCell(7).setText(entity.getFinInstitutionsName() != null ? entity.getFinInstitutionsName() : "");
                        dataRow.getCell(8).setText(entity.getTotalCountOfCredits() != null ? entity.getTotalCountOfCredits().toString() : "");
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
                    creteTitle(doc,"Сведения по аморальному образу жизни");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по аморальному образу жизни", Arrays.asList("№", "Орган выявивший", "Гражданство", "Дата решения", "Сумма штрафа"));

                    
                    int number = 1;
                    for (ImmoralLifestyle r : result.getAmoral()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getAuthority_detected() != null ? r.getAuthority_detected() : "");
                        dataRow.getCell(2).setText(r.getCitizenship_id() != null ? r.getCitizenship_id() : "");
                        dataRow.getCell(3).setText(r.getDecision_date() != null ? r.getDecision_date().toString() : "");
                        dataRow.getCell(4).setText(r.getFine_amount() != null ? r.getFine_amount().toString() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                    // Save the document as needed
                }
            } catch (Exception e) {
                System.out.println("Exception while adding immoral lifestyle entities table: " + e.getMessage());
            }try {
                if (result.getMzEntities() != null && !result.getMzEntities().isEmpty()) {
                    creteTitle(doc,"Сведения по МЗ");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по МЗ", Arrays.asList("№", "Код болезни", "Регистрация", "Статус МЗ", "Медицинская организация"));

                    
                    int number = 1;
                    for (MzEntity r : result.getMzEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getDiseaseCode() != null ? r.getDiseaseCode() : "");
                        dataRow.getCell(2).setText(r.getReg() != null ? r.getReg() : "");
                        dataRow.getCell(3).setText(r.getStatusMz() != null ? r.getStatusMz() : "");
                        dataRow.getCell(4).setText(r.getMedicalOrg() != null ? r.getMedicalOrg() : "");
                        number++;                  }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding MZ entities table: " + e.getMessage());
            }try {
                if (result.getWantedListEntities() != null && !result.getWantedListEntities().isEmpty()) {
                    creteTitle(doc,"Сведения по разыскиваемым");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения по разыскиваемым", Arrays.asList("№", "Дни", "Орган", "Статус", "Дата актуальности"));

                    
                    int number = 1;
                    for (WantedListEntity r : result.getWantedListEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(r.getDays() != null ? r.getDays().toString() : "");
                        dataRow.getCell(2).setText(r.getOrgan() != null ? r.getOrgan() : "");
                        dataRow.getCell(3).setText(r.getStatus() != null ? r.getStatus() : "");
                        dataRow.getCell(4).setText(r.getRelevanceDate() != null ? r.getRelevanceDate().toString() : "");
                        number++;                    }
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
                    creteTitle(doc, "Сведения о юридическом лице");
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    makeTableByProperties(doc, table, "Сведения о юридическом лице", Arrays.asList(
                            "БИН",
                            "Наименование организации",
                            "Наименование ОКЭД",
                            "Статус ЮЛ"
                    ));

                    int number = 1;
                    for (MvUl a : result.getMvUls()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getBin() != null ? a.getBin() : "");
                        dataRow.getCell(2).setText(a.getFull_name_kaz() != null ? a.getFull_name_kaz() : "");
                        dataRow.getCell(3).setText(a.getHead_organization() != null ? a.getHead_organization() : "");
                        dataRow.getCell(4).setText(a.getUl_status() != null ? a.getUl_status() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Exception while adding MV UL table: " + e.getMessage());
            }

            try {
                if (result.getMvUlFounderFls() != null && !result.getMvUlFounderFls().isEmpty()) {
                    creteTitle(doc, "Сведения об участниках ЮЛ");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        String name = mvUlRepo.getNameByBin(a.getBin_org());
                        dataRow.getCell(1).setText(name != null ? name : "Нет");
                        dataRow.getCell(2).setText(a.getReg_date() != null ? a.getReg_date().toString() : "Нет даты");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding MV Ul Founder Fl table: " + e.getMessage());
            }

            try {
                if (result.getAccountantListEntities() != null && !result.getAccountantListEntities().isEmpty()) {
                    creteTitle(doc, "Список бухгалтеров");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Список бухгалтеров", Arrays.asList(
                            "№",
                            "ИИН",
                            "Проф.",
                            "Фамилия",
                            "Имя"
                    ));

                    int number = 1;
                    for (AccountantListEntity a : result.getAccountantListEntities()) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getIin() != null ? a.getIin() : "");
                        dataRow.getCell(2).setText(a.getProf() != null ? a.getProf() : "");
                        dataRow.getCell(3).setText(a.getLname() != null ? a.getLname() : "");
                        dataRow.getCell(4).setText(a.getFname() != null ? a.getFname() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding AccountantListEntities table: " + e.getMessage());
            }

            try {
                List<Omn> omns = result.getOmns();
                if (omns != null && !omns.isEmpty()) {
                    creteTitle(doc, "ОМНС");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getRnn() != null ? a.getRnn() : "");
                        dataRow.getCell(2).setText(a.getTaxpayer_name() != null ? a.getTaxpayer_name() : "");
                        dataRow.getCell(3).setText(a.getTaxpayer_fio() != null ? a.getTaxpayer_fio() : "");
                        dataRow.getCell(4).setText(a.getLeader_fio() != null ? a.getLeader_fio() : "");
                        dataRow.getCell(5).setText(a.getLeader_iin() != null ? a.getLeader_iin() : "");
                        dataRow.getCell(6).setText(a.getLeader_rnn() != null ? a.getLeader_rnn() : "");
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
                    creteTitle(doc, "Транспорт");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Транспорт", Arrays.asList(
                            "№",
                            "Адрес",
                            "Гос. Номер",
                            "Номер серии рег.",
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getOwner_address() != null ? a.getOwner_address() : "");
                        dataRow.getCell(2).setText(a.getGov_number() != null ? a.getGov_number() : "");
                        dataRow.getCell(3).setText(a.getReg_series_num() != null ? a.getReg_series_num() : "");
                        dataRow.getCell(4).setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.getCell(5).setText(a.getReg_reason() != null ? a.getReg_reason() : "");
                        dataRow.getCell(6).setText(a.getVin() != null ? a.getVin() : "");
                        dataRow.getCell(7).setText(a.getEquipment_spec() != null ? a.getEquipment_spec() : "");
                        dataRow.getCell(8).setText(a.getEquipment_type() != null ? a.getEquipment_type() : "");
                        dataRow.getCell(9).setText(a.getEquipment_form() != null ? a.getEquipment_form() : "");
                        dataRow.getCell(10).setText(a.getBrand() != null ? a.getBrand() : "");
                        dataRow.getCell(11).setText(a.getEquipment_model() != null ? a.getEquipment_model() : "");
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
                    creteTitle(doc, "МШЭС");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "МШЭС", Arrays.asList(
                            "№",
                            "Тип оборудования",
                            "Модель оборудования",
                            "VIN",
                            "Гос. номер",
                            "Дата регистрации"
                    ));

                    int number = 1;
                    for (Msh a : mshes) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(2).setText(a.getEquipmentModel() != null ? a.getEquipmentModel() : "");
                        dataRow.getCell(3).setText(a.getVin() != null ? a.getVin() : "");
                        dataRow.getCell(4).setText(a.getGovNumber() != null ? a.getGovNumber() : "");
                        dataRow.getCell(5).setText(a.getRegDate() != null ? a.getRegDate().toString() : "");
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
                    creteTitle(doc,"Дорманс");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getRnn() != null ? a.getRnn() : "");
                        dataRow.getCell(2).setText(a.getTaxpayer_name() != null ? a.getTaxpayer_name() : "");
                        dataRow.getCell(3).setText(a.getTaxpayer_fio() != null ? a.getTaxpayer_fio() : "");
                        dataRow.getCell(4).setText(a.getLeader_fio() != null ? a.getLeader_fio() : "");
                        dataRow.getCell(5).setText(a.getLeader_iin() != null ? a.getLeader_iin() : "");
                        dataRow.getCell(6).setText(a.getLeader_rnn() != null ? a.getLeader_rnn() : "");
                        dataRow.getCell(7).setText(a.getOrder_date() != null ? a.getOrder_date() : "");
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
                    creteTitle(doc, "Банкроты");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Банкроты", Arrays.asList(
                            "№",
                            "Документ",
                            "Дата обновления",
                            "Причина"
                    ));

                    int number = 1;
                    for (Bankrot a : bankrots) {
                        XWPFTableRow dataRow = table.getRow(number);
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getDocument() != null ? a.getDocument() : "");
                        try {
                            dataRow.getCell(2).setText(a.getUpdate_dt() != null ? a.getUpdate_dt().toString() : "Дата отсутствует");
                        } catch (Exception e) {
                            dataRow.getCell(2).setText("Дата отсутствует");
                        }
                        dataRow.getCell(3).setText(a.getReason() != null ? a.getReason() : "");
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
                    creteTitle(doc, "Администрация");
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
                        XWPFTableRow dataRow = table.getRow(number);
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getMaterial_num() != null ? a.getMaterial_num() : "");
                        dataRow.getCell(2).setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.getCell(3).setText(a.getFifteen() != null ? a.getFifteen() : "");
                        dataRow.getCell(4).setText(a.getSixteen() != null ? a.getSixteen() : "");
                        dataRow.getCell(5).setText(a.getSeventeen() != null ? a.getSeventeen() : "");
                        dataRow.getCell(6).setText(a.getUl_org_name() != null ? a.getUl_org_name() : "");
                        dataRow.getCell(7).setText(a.getUl_adress() != null ? a.getUl_adress() : "");
                        dataRow.getCell(8).setText(a.getVehicle_brand() != null ? a.getVehicle_brand() : "");
                        dataRow.getCell(9).setText(a.getState_auto_num() != null ? a.getState_auto_num() : "");
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
                    creteTitle(doc, "Преступления");
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
                        XWPFTableRow dataRow = table.getRow(number);
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getCourt_name() != null ? a.getCourt_name() : "");
                        dataRow.getCell(2).setText(a.getCourt_dt() != null ? a.getCourt_dt() : "");
                        dataRow.getCell(3).setText(a.getDecision() != null ? a.getDecision() : "");
                        dataRow.getCell(4).setText(a.getCrime_name() != null ? a.getCrime_name() : "");
                        dataRow.getCell(5).setText(a.getSentence() != null ? a.getSentence() : "");
                        dataRow.getCell(6).setText(a.getAdd_info() != null ? a.getAdd_info() : "");
                        dataRow.getCell(7).setText(a.getTreatment() != null ? a.getTreatment() : "");
                        dataRow.getCell(8).setText(a.getErdr() != null ? a.getErdr() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding Criminals table: " + e.getMessage());
            }

            try {
                List<BlockEsf> blockEsfs = result.getBlockEsfs();
                if (blockEsfs != null && !blockEsfs.isEmpty()) {
                    creteTitle(doc, "Блокировка ЕСФ");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Блокировка ЕСФ", Arrays.asList(
                            "№",
                            "Дата начала",
                            "Дата окончания",
                            "Дата обновления"
                    ));

                    int number = 1;
                    for (BlockEsf a : blockEsfs) {
                        XWPFTableRow dataRow = table.getRow(number);
                        dataRow.getCell(0).setText(String.valueOf(number));
                        try {
                            dataRow.getCell(1).setText(a.getStart_dt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(1).setText("Нет");
                        }
                        try {
                            dataRow.getCell(2).setText(a.getEnd_dt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(2).setText("Нет");
                        }
                        try {
                            dataRow.getCell(3).setText(a.getUpdate_dt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(3).setText("Нет");
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
                    creteTitle(doc, "Объекты НДС");
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
                        XWPFTableRow dataRow = table.getRow(number);
                        dataRow.getCell(0).setText(String.valueOf(number));
                        try {
                            dataRow.getCell(1).setText(a.getStartDt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(1).setText("Нет");
                        }
                        try {
                            dataRow.getCell(2).setText(a.getEndDt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(2).setText("Нет");
                        }
                        dataRow.getCell(3).setText(a.getReason() != null ? a.getReason() : "");
                        try {
                            dataRow.getCell(4).setText(a.getUpdateDt().toString());
                        } catch (Exception e) {
                            dataRow.getCell(4).setText("Нет");
                        }
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding NdsEntity table: " + e.getMessage());
            }

//            try {
//                List<MvRnOld> mvRnOlds = result.getMvRnOlds();
//                if (mvRnOlds != null && !mvRnOlds.isEmpty()) {
//                    creteTitle(doc, "Прежний адрес прописки");
//                    XWPFTable table = doc.createTable();
//                    makeTableByProperties(doc, table, "Прежний адрес прописки", Arrays.asList(
//                            "№",
//                            "Назначение использования",
//                            "Статус недвижимости",
//                            "Адрес",
//                            "История адресов",
//                            "Тип собственности",
//                            "Вид собственности",
//                            "Статус характеристики недвижимости",
//                            "Дата регистрации в реестре",
//                            "Дата окончания регистрации",
//                            "Возникновение права в реестре",
//                            "Статус в реестре"
//                    ));
//
//                    int number = 1;
//                    for (MvRnOld a : mvRnOlds) {
//                        XWPFTableRow dataRow = table.getRow(number);
//                        dataRow.getCell(0).setText(String.valueOf(number));
//                        dataRow.getCell(1).setText(a.getType_of_property_rus() != null ? a.getProperty_type_rus() : "");
//                        dataRow.getCell(2).setText(a.getEstate_status_rus() != null ? a.getProperty_status() : "");
//                        dataRow.getCell(3).setText(a.getAddress() != null ? a.getAddress() : "");
//                        dataRow.getCell(4).setText(a.getAddress_history() != null ? a.getAddress_history() : "");
//                        dataRow.getCell(5).setText(a.getProperty_type() != null ? a.getProperty_type() : "");
//                        dataRow.getCell(6).setText(a.getOwnership_type() != null ? a.getOwnership_type() : "");
//                        dataRow.getCell(7).setText(a.getProperty_status() != null ? a.getProperty_status() : "");
//                        dataRow.getCell(8).setText(a.getReg_date() != null ? a.getReg_date().toString() : "");
//                        dataRow.getCell(9).setText(a.getEnd_reg_date() != null ? a.getEnd_reg_date().toString() : "");
//                        dataRow.getCell(10).setText(a.getRight_occur_date() != null ? a.getRight_occur_date().toString() : "");
//                        dataRow.getCell(11).setText(a.getReg_status() != null ? a.getReg_status() : "");
//                        number++;
//                    }
//                    setMarginBetweenTables(doc);
//                }
//            } catch (Exception e) {
//                System.out.println("Error adding MvRnOld table: " + e.getMessage());
//            }
            try {
                List<FpgTempEntity> fpgTempEntities = result.getFpgTempEntities();
                if (fpgTempEntities != null && !fpgTempEntities.isEmpty()) {
                    creteTitle(doc,"Временные объекты ФПГ");
                    XWPFTable table = doc.createTable();
                    makeTableByProperties(doc, table, "Временные объекты ФПГ", Arrays.asList(
                            "№",
                            "Бенефициар"
                    ));

                    int number = 1;
                    for (FpgTempEntity a : fpgTempEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getBeneficiary() != null ? a.getBeneficiary() : "");
                        number++;
                    }
                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding FpgTempEntity table: " + e.getMessage());
            }try {
                List<Pdl> pdls = result.getPdls();
                if (pdls != null && !pdls.isEmpty()) {
                    creteTitle(doc,"ПДЛ");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getIin() != null ? a.getIin() : "");
                        dataRow.getCell(2).setText(a.getOrganization_fullname() != null ? a.getOrganization_fullname() : "");
                        dataRow.getCell(3).setText(a.getFio() != null ? a.getFio() : "");
                        dataRow.getCell(4).setText(a.getOrgan() != null ? a.getOrgan() : "");
                        dataRow.getCell(5).setText(a.getOblast() != null ? a.getOblast() : "");
                        dataRow.getCell(6).setText(a.getSpouse_fio() != null ? a.getSpouse_fio() : "");
                        dataRow.getCell(7).setText(a.getSpouse_organ() != null ? a.getSpouse_organ() : "");
                        dataRow.getCell(8).setText(a.getSpouse_position() != null ? a.getSpouse_position() : "");
                        dataRow.getCell(9).setText(a.getSpouse_iin() != null ? a.getSpouse_iin() : "");
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
                    creteTitle(doc,"Производители товаров");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getSspName() != null ? a.getSspName() : "");
                        dataRow.getCell(2).setText(String.valueOf(a.getCount()));
                        dataRow.getCell(3).setText(a.getProducer() != null ? a.getProducer() : "");
                        dataRow.getCell(4).setText(a.getStatus() != null ? a.getStatus() : "");
                        dataRow.getCell(5).setText(a.getRegion() != null ? a.getRegion() : "");
                        dataRow.getCell(6).setText(a.getSztp() != null ? a.getSztp() : "");
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
                    creteTitle(doc,"Адрес");
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
                        dataRow.getCell(0).setText(regAddressUlEntity.getRegDate().toString());
                    } catch (Exception e) {
                        dataRow.getCell(0).setText("");
                    }
                    dataRow.getCell(1).setText(regAddressUlEntity.getOrgNameRu() != null ? regAddressUlEntity.getOrgNameRu() : "");
                    dataRow.getCell(2).setText(regAddressUlEntity.getRegAddrRegionRu() != null ? regAddressUlEntity.getRegAddrRegionRu() : "");
                    dataRow.getCell(3).setText(regAddressUlEntity.getRegAddrDistrictRu() != null ? regAddressUlEntity.getRegAddrDistrictRu() : "");
                    dataRow.getCell(4).setText(regAddressUlEntity.getRegAddrRuralDistrictRu() != null ? regAddressUlEntity.getRegAddrRuralDistrictRu() : "");
                    dataRow.getCell(5).setText(regAddressUlEntity.getRegAddrLocalityRu() != null ? regAddressUlEntity.getRegAddrLocalityRu() : "");
                    dataRow.getCell(6).setText(regAddressUlEntity.getRegAddrStreetRu() != null ? regAddressUlEntity.getRegAddrStreetRu() : "");
                    dataRow.getCell(7).setText(regAddressUlEntity.getRegAddrBuildingNum() != null ? regAddressUlEntity.getRegAddrBuildingNum() : "");
                    dataRow.getCell(8).setText(regAddressUlEntity.getRegAddrBlockNum() != null ? regAddressUlEntity.getRegAddrBlockNum() : "");
                    dataRow.getCell(9).setText(regAddressUlEntity.getRegAddrBuildingBodyNum() != null ? regAddressUlEntity.getRegAddrBuildingBodyNum() : "");
                    dataRow.getCell(10).setText(regAddressUlEntity.getRegAddrOffice() != null ? regAddressUlEntity.getRegAddrOffice() : "");
                    dataRow.getCell(11).setText(regAddressUlEntity.getOkedNameRu() != null ? regAddressUlEntity.getOkedNameRu() : "");
                    dataRow.getCell(12).setText(regAddressUlEntity.getUl_status() != null ? regAddressUlEntity.getUl_status() : "");
                    dataRow.getCell(13).setText(regAddressUlEntity.getActive() ? "Активен" : "Неактивен");

                    setMarginBetweenTables(doc);
                }
            } catch (Exception e) {
                System.out.println("Error adding RegAddressUlEntity table: " + e.getMessage());
            }

            try {
                List<SvedenyaObUchastnikovUlEntity> svedenyaObUchastnikovUlEntities = result.getSvedenyaObUchastnikovUlEntities();
                if (svedenyaObUchastnikovUlEntities != null && !svedenyaObUchastnikovUlEntities.isEmpty()) {
                    creteTitle(doc,"Сведения об участниках ЮЛ");
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
                        dataRow.getCell(0).setText(String.valueOf(number));
                        dataRow.getCell(1).setText(a.getFIOorUlName() != null ? a.getFIOorUlName() : "");
                        dataRow.getCell(2).setText(a.getIdentificator() != null ? a.getIdentificator() : "");
                        dataRow.getCell(3).setText(a.getReg_date() != null ? a.getReg_date() : "");
                        dataRow.getCell(4).setText(a.getRisk() != null ? a.getRisk() : "");
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