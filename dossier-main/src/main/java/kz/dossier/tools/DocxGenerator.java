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

    public void generateDoc(NodesFL result, ByteArrayOutputStream baos) throws IOException, InvalidFormatException {
        try (XWPFDocument doc = new XWPFDocument()) {

            try {
                if (result.getMvFls() != null || result.getMvFls().size() < 0) {
                    XWPFTable table1 = doc.createTable();
                    table1.setWidth("100%");

                    XWPFTableRow row3 = table1.createRow();
                    XWPFTableCell cell = row3.addNewTableCell();

                    cell.setWidth("100%");
                    cell.setColor("808080");
                    XWPFParagraph paragraph1 = cell.addParagraph();
                    paragraph1.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run = paragraph1.createRun();
                    run.setText("Сведения о физическом лице");

                    XWPFTableRow row = table1.createRow();
                    row.addNewTableCell().setText("Фото");
                    row.addNewTableCell().setText("ИИН");
                    row.addNewTableCell().setText("ФИО");
                    row.addNewTableCell().setText("Резидент");
                    row.addNewTableCell().setText("Национальность");
                    row.addNewTableCell().setText("Дата смерти");

                    XWPFTableRow row1 = table1.createRow();
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
                    XWPFParagraph paragraph = doc.createParagraph();
                    XWPFRun run2 = paragraph.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ

                }
            } catch (Exception e) {
                System.out.println("Mv_Fl table add exception");
            }
            try {
                if (result.getMvRnOlds() != null || result.getMvRnOlds().size() < 0) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    XWPFTableRow row = table.createRow();
                    XWPFTableCell cell = row.addNewTableCell();

                    cell.setWidth("100%");
                    cell.setColor("808080");
                    XWPFParagraph paragraph = cell.addParagraph();
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run = paragraph.createRun();
                    run.setText("Адреса прописки");

                    XWPFTableRow row1 = table.createRow();
                    row1.addNewTableCell().setText("Страна");
                    row1.addNewTableCell().setText("Город");
                    row1.addNewTableCell().setText("Адрес");
                    row1.addNewTableCell().setText("Регион");
                    row1.addNewTableCell().setText("Дата прописки");

                    XWPFTableRow row2 = table.createRow();

                    row2.addNewTableCell().setText(result.getMvRnOlds().get(0).getAddress_history_kaz());
                    row2.addNewTableCell().setText(result.getMvRnOlds().get(0).getRegister_emergence_rights_rus());
                    row2.addNewTableCell().setText(result.getMvRnOlds().get(0).getAddress_history_kaz());
                    row2.addNewTableCell().setText(result.getMvRnOlds().get(0).getRegister_end_date());
                    row2.addNewTableCell().setText(result.getMvRnOlds().get(0).getAddress_history_kaz());
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ

                }
            } catch (Exception e) {
                System.out.println("MV_Rn_Old table add exception");
            }
            try {
                List<MvIinDoc> docs = result.getMvIinDocs();
                if (docs != null && !docs.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    XWPFTableRow row = table.createRow();
                    XWPFTableCell cell = row.addNewTableCell();

                    cell.setWidth("100%");
                    cell.setColor("808080");
                    XWPFParagraph paragraph = cell.addParagraph();
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run = paragraph.createRun();
                    run.setText("Документы");
                    // Add the column headers
                    String[] headers = {"Типа Документа", "Орган выдачи", "Дата выдачи", "Срок до", "Номер документа"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell1 = columnHeaderRow.addNewTableCell();
                        cell1.setText(header);
                        XWPFParagraph paragraph1 = cell1.getParagraphs().get(0);
                        paragraph1.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run1 = paragraph.createRun();
                        run1.setBold(true);
                        cell1.setColor("D3D3D3"); // Light gray color
                    }

                    // Add the data rows
                    for (MvIinDoc doci : docs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(doci.getDoc_type_ru_name());
                        dataRow.addNewTableCell().setText(doci.getIssue_organization_ru_name());
                        dataRow.addNewTableCell().setText(doci.getIssue_date().toString());
                        dataRow.addNewTableCell().setText(doci.getExpiry_date().toString());
                        dataRow.addNewTableCell().setText(doci.getDoc_type_ru_name());
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("MV_Iin_Doc table add exception");
            }
            try {
                List<School> schools = result.getSchools();
                if (schools != null && !schools.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Школы");
                    headingRun.setBold(true);

                    // Add the column headers
                    String[] headers = {"БИН", "Название школы", "Класс", "Дата поступления", "Дата окончания"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                    }

                    // Add the data rows
                    for (School school : schools) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(school.getSchool_code());
                        dataRow.addNewTableCell().setText(school.getSchool_name());
                        dataRow.addNewTableCell().setText(school.getGrade());
                        dataRow.addNewTableCell().setText(school.getStart_date().toString());
                        dataRow.addNewTableCell().setText(school.getEnd_date().toString());
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding school table: " + e.getMessage());
            }
            try {
                List<Universities> universities = result.getUniversities();
                if (universities != null && !universities.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Вузы");
                    headingRun.setBold(true);

                    // Add the column headers
                    String[] headers = {"БИН", "Название вуза", "Специализация", "Дата поступления", "Дата окончания", "Длительность обучения", "Курс"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding university table: " + e.getMessage());
            }
            try {
                List<MvAutoFl> autos = result.getMvAutoFls();
                if (autos != null && !autos.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Транспорт");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    headingCell.setColor("D3D3D3"); // Light gray color

                    // Add the column headers
                    String[] headers = {"№", "Статус", "Регистрационный номер", "Марка модель", "Дата выдачи свидетельства", "Дата снятия", "Год выпуска", "Категория", "VIN/Кузов/Шосси", "Серия"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
                    int number = 1;
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
                    for (MvAutoFl auto : autos) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));

                        try {
                            if (formatter.format(auto.getEnd_date()).compareTo(formatter.format(new java.util.Date())) > 0) {
                                dataRow.addNewTableCell().setText("Действителен");
                            } else {
                                dataRow.addNewTableCell().setText("Не действителен");
                            }
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }

                        try {
                            dataRow.addNewTableCell().setText(auto.getReg_number());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }

                        try {
                            dataRow.addNewTableCell().setText(auto.getBrand_model());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }

                        try {
                            dataRow.addNewTableCell().setText(auto.getDate_certificate().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }

                        try {
                            dataRow.addNewTableCell().setText(auto.getEnd_date().toString());
                        } catch (Exception e) {
                            dataRow.addNewTableCell().setText("");
                        }

                        dataRow.addNewTableCell().setText(auto.getRelease_year_tc());
                        dataRow.addNewTableCell().setText(auto.getOwner_category());
                        dataRow.addNewTableCell().setText(auto.getVin_kuzov_shassi());
                        dataRow.addNewTableCell().setText(auto.getSeries_reg_number());
                        number++;
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding auto table: " + e.getMessage());
            }
            List<FlRelativiesDTO> fl_relatives = result.getFl_relatives();
            try {
                if (fl_relatives != null && !fl_relatives.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");
                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Родственные связи");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    headingCell.setColor("D3D3D3"); // Light gray color

                    // Add the column headers
                    String[] headers = {"№", "Статус по отношению к родственнику", "ФИО", "ИИН", "Дата рождения", "Дата регистрации брака", "Дата расторжения брака"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding relatives table: " + e.getMessage());
            }
            try {
                List<FlContacts> contacts = result.getContacts();
                if (contacts != null && !contacts.isEmpty()) {
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    XWPFTableRow row = table.createRow();
                    XWPFTableCell cell1 = row.addNewTableCell();

                    cell1.setWidth("100%");
                    cell1.setColor("808080");
                    XWPFParagraph paragraph1 = cell1.addParagraph();
                    paragraph1.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run1 = paragraph1.createRun();
                    run1.setText("Контактные данные ФЛ");
                    // Add the column headers
                    String[] headers = {"№", "Телефон", "Почта", "Источник"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
                    int number = 1;
                    for (FlContacts contact : contacts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(contact.getPhone());
                        dataRow.addNewTableCell().setText(contact.getEmail());
                        dataRow.addNewTableCell().setText(contact.getSource());
                        number++;
                    }
                    XWPFParagraph paragraph2 = doc.createParagraph();
                    XWPFRun run2 = paragraph2.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding contacts table: " + e.getMessage());
            }try {
                List<MillitaryAccount> militaryAccounts = result.getMillitaryAccounts();
                if (militaryAccounts != null && !militaryAccounts.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Воинский учет");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    headingCell.setColor("D3D3D3"); // Light gray color

                    // Add the column headers
                    String[] headers = {"№", "БИН воинской части", "Дата службы с", "Дата службы по"};
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
                    int number = 1;
                    for (MillitaryAccount account : militaryAccounts) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(account.getBin());
                        dataRow.addNewTableCell().setText(account.getDate_start());
                        dataRow.addNewTableCell().setText(account.getDate_end());
                        number++;
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding military accounts table: " + e.getMessage());
            }try {
                List<ConvictsJustified> convictsJustifieds = result.getConvictsJustifieds();
                if (convictsJustifieds != null && !convictsJustifieds.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Наименование риска: \"Осужденные\" Количество найденных инф: " + convictsJustifieds.size());
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    headingCell.setColor("D3D3D3"); // Light gray color

                    // Add the column headers
                    String[] headers = {
                            "№",
                            "Дата рассмотрения в суде 1 инстанции",
                            "Суд 1 инстанции",
                            "Решение по лицу",
                            "Мера наказания по договору",
                            "Квалификация"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding convicts justified table: " + e.getMessage());
            }try {
                List<ConvictsTerminatedByRehab> convictsTerminatedByRehabs = result.getConvictsTerminatedByRehabs();
                if (convictsTerminatedByRehabs != null && !convictsTerminatedByRehabs.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Административные штрафы");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    headingCell.setColor("D3D3D3"); // Light gray color

                    // Add the column headers
                    String[] headers = {
                            "№",
                            "Орган выявивший правонарушение",
                            "Дата заведения",
                            "Квалификация",
                            "Решение",
                            "Уровень тяжести"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding convicts terminated by rehab table: " + e.getMessage());
            }try {
                List<BlockEsf> blockEsfs = result.getBlockEsfs();
                if (blockEsfs != null && !blockEsfs.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setWidth("100%");
                    headingCell.setColor("808080");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run1 = headingParagraph.createRun();
                    run1.setText("Блокировка ЭСФ");


                    // Add the column headers
                    String[] headers = {
                            "№",
                            "Дата блокировки",
                            "Дата востановления",
                            "Дата обновления"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
                    int number = 1;
                    for (BlockEsf r : blockEsfs) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));

                        dataRow.addNewTableCell().setText(r.getStart_dt() != null ? r.getStart_dt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getEnd_dt() != null ? r.getEnd_dt().toString() : "");
                        dataRow.addNewTableCell().setText(r.getUpdate_dt() != null ? r.getUpdate_dt().toString() : "");
                        number++;
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding block esf table: " + e.getMessage());
            }try {
                List<MvUlFounderFl> mvUlFounderFls = result.getMvUlFounderFls();
                if (mvUlFounderFls != null && !mvUlFounderFls.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    headingCell.setWidth("100%");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Сведения об участниках ЮЛ");

                    // Add the column headers
                    String[] headers = {
                            "№",
                            "БИН",
                            "Наименование ЮЛ",
                            "Дата регистрации"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding mv ul founder fl table: " + e.getMessage());
            }try {
                List<NdsEntity> ndsEntities = result.getNdsEntities();
                if (ndsEntities != null && !ndsEntities.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    headingCell.setWidth("100%");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("НДС");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    // Add the column headers
                    String[] headers = {
                            "№",
                            "Дата начала",
                            "Дата конца",
                            "Дата обновления",
                            "Причина"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
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
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
                }
            } catch (Exception e) {
                System.out.println("Exception while adding nds entities table: " + e.getMessage());
            }    try {
                List<IpgoEmailEntity> ipgoEmailEntities = result.getIpgoEmailEntities();
                if (ipgoEmailEntities != null && !ipgoEmailEntities.isEmpty()) {
                    // Create the table
                    XWPFTable table = doc.createTable();
                    table.setWidth("100%");

                    // Create the heading row
                    XWPFTableRow headingRow = table.createRow();
                    XWPFTableCell headingCell = headingRow.addNewTableCell();
                    headingCell.setColor("808080");
                    headingCell.setWidth("100%");
                    XWPFParagraph headingParagraph = headingCell.addParagraph();
                    headingParagraph.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun headingRun = headingParagraph.createRun();
                    headingRun.setText("Сведения по ИПГО");
                    headingRun.setBold(true);
                    headingCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    // Add the column headers
                    String[] headers = {
                            "№",
                            "Департамент",
                            "Должность",
                            "ИПГО почта"
                    };
                    XWPFTableRow columnHeaderRow = table.createRow();
                    for (String header : headers) {
                        XWPFTableCell cell = columnHeaderRow.addNewTableCell();
                        XWPFParagraph paragraph = cell.addParagraph();
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText(header);
                        run.setBold(true);
                        cell.setColor("D3D3D3"); // Light gray color
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    }

                    // Add the data rows
                    int number = 1;
                    for (IpgoEmailEntity r : ipgoEmailEntities) {
                        XWPFTableRow dataRow = table.createRow();
                        dataRow.addNewTableCell().setText(String.valueOf(number));
                        dataRow.addNewTableCell().setText(r.getOrgan() != null ? r.getOrgan().toString() : "");
                        dataRow.addNewTableCell().setText(r.getPosition() != null ? r.getPosition() : "");
                        dataRow.addNewTableCell().setText(r.getEmail() != null ? r.getEmail().toString() : "");
                        number++;
                    }
                    XWPFParagraph paragraph1 = doc.createParagraph();
                    XWPFRun run2 = paragraph1.createRun();
                    run2.addBreak();  // Добавляем перенос строки
                    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ
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


//    XWPFParagraph paragraph1 = doc.createParagraph();
//    XWPFRun run2 = paragraph1.createRun();
//    run2.addBreak();  // Добавляем перенос строки
//    run2.setText(" "); // Добавляем пробел, чтобы создать визуальный отступ


    public void generate(NodesUL result, ByteArrayOutputStream baos) throws IOException {
        XWPFDocument document = new XWPFDocument();

        // Add "Сведения о юридическом лице" section
        List<MvUl> mvUl = result.getMvUls();
        if (mvUl != null && !mvUl.isEmpty()) {
            addTableToDocument(document, "Сведения о юридическом лице", mvUl, new String[]{"БИН", "Наименование организаци", "Наименование ОКЭД", "Статус ЮЛ"});
        }

        // Add "Сведения об участниках ЮЛ" section
        List<MvUlFounderFl> mvUlFounderFls = result.getMvUlFounderFls();
        if (mvUlFounderFls != null && !mvUlFounderFls.isEmpty()) {
            addTableToDocument(document, "Сведения об участниках ЮЛ", mvUlFounderFls, new String[]{"№", "БИН", "Наименование ЮЛ", "Дата регистрации"});
        }

        // Add "Список бухгалтеров" section
        List<AccountantListEntity> accountantListEntities = result.getAccountantListEntities();
        if (accountantListEntities != null && !accountantListEntities.isEmpty()) {
            addTableToDocument(document, "Список бухгалтеров", accountantListEntities, new String[]{"ИИН", "Проф.", "Фамилия", "Имя"});
        }

        // Add "ОМНС" section
        List<Omn> omns = result.getOmns();
        if (omns != null && !omns.isEmpty()) {
            addTableToDocument(document, "ОМНС", omns, new String[]{"РНН", "Название налогоплательщика", "ФИО налогоплательщика", "ФИО руководителя", "ИИН руководителя", "РНН руководителя"});
        }

        // Add "Транспорт" section
//        List<Equipment> equipmentList = result.getEquipment();
//        if (equipmentList != null && !equipmentList.isEmpty()) {
//            addTableToDocument(document, "Транспорт", equipmentList, new String[]{"Адрес", "Гос. Номер", "Номер серии рег.", "Дата регистрации", "Причина", "VIN", "Спец.", "Тип", "Форма", "Брэнд", "Модель"});
//        }
//
//        // Add "МШЭС" section
//        List<Msh> mshes = result.getMshes();
//        if (mshes != null && !mshes.isEmpty()) {
//            addTableToDocument(document, "МШЭС", mshes, new String[]{"Тип оборудования", "Модель оборудования", "VIN", "Гос. номер", "Дата регистрации"});
//        }
//
//        // Add "Дорманс" section
//        List<Dormant> dormans = result.getDormants();
//        if (dormans != null && !dormans.isEmpty()) {
//            addTableToDocument(document, "Дорманс", dormans, new String[]{"РНН", "Название налогоплательщика", "ФИО налогоплательщика", "ФИО руководителя", "ИИН руководителя", "РНН руководителя", "Дата заказа"});
//        }
//
//        // Add "Банкроты" section
//        List<Bankrot> bankrots = result.getBankrots();
//        if (bankrots != null && !bankrots.isEmpty()) {
//            addTableToDocument(document, "Банкроты", bankrots, new String[]{"Документ", "Дата обновления", "Причина"});
//        }
//
//        // Add "Администрация" section
//        List<Adm> adms = result.getAdms();
//        if (adms != null && !adms.isEmpty()) {
//            addTableToDocument(document, "Администрация", adms, new String[]{"Номер материала", "Дата регистрации", "15", "16", "17", "Наименование юр. лица", "Адрес юр. лица", "Марка автомобиля", "Гос. Номер авто"});
//        }
//
//        // Add "Преступления" section
//        List<Criminals> criminals = result.getCriminals();
//        if (criminals != null && !criminals.isEmpty()) {
//            addTableToDocument(document, "Преступления", criminals, new String[]{"Наименование суда", "Дата судебного решения", "Решение", "Название преступления", "Приговор", "Дополнительная информация", "Обращение", "ЕРДР"});
//        }
//
//        // Add "Блокировка ЕСФ" section
//        List<BlockEsf> blockEsfs = result.getBlockEsfs();
//        if (blockEsfs != null && !blockEsfs.isEmpty()) {
//            addTableToDocument(document, "Блокировка ЕСФ", blockEsfs, new String[]{"Дата начала", "Дата окончания", "Дата обновления"});
//        }
//
//        // Add "Объекты НДС" section
//        List<NdsEntity> ndsEntities = result.getNdsEntities();
//        if (ndsEntities != null && !ndsEntities.isEmpty()) {
//            addTableToDocument(document, "Объекты НДС", ndsEntities, new String[]{"Дата начала", "Дата окончания", "Причина", "Дата обновления"});
//        }
//
//        // Add "mv_rn_old" section
//        List<MvRnOld> mvRnOlds = result.getMvRnOlds();
//        if (mvRnOlds != null && !mvRnOlds.isEmpty()) {
//            addTableToDocument(document, "mv_rn_old", mvRnOlds, new String[]{"Назначение использования", "Статус недвижимости", "Адрес", "История адресов", "Тип собственности", "Вид собственности", "Статус характеристики недвижимости", "Дата регистрации в реестре", "Дата окончания регистрации", "Возникновение права в реестре", "Статус в реестре"});
//        }
//
//        // Add "Временные объекты ФПГ" section
//        List<FpgTempEntity> fpgTempEntities = result.getFpgTempEntities();
//        if (fpgTempEntities != null && !fpgTempEntities.isEmpty()) {
//            addTableToDocument(document, "Временные объекты ФПГ", fpgTempEntities, new String[]{"№", "Бенефициар"});
//        }
//
//        // Add "ПДЛ" section
//        List<Pdl> pdls = result.getPdls();
//        if (pdls != null && !pdls.isEmpty()) {
//            addTableToDocument(document, "ПДЛ", pdls, new String[]{"ИИН", "Полное наименование организации", "ФИО", "Орган", "Область", "ФИО супруг(и)", "Орган супруг(и)", "Должность супруга", "ИИН супруга"});
//        }
//
//        // Add "Производители товаров" section
//        List<CommodityProducer> commodityProducers = result.getCommodityProducers();
//        if (commodityProducers != null && !commodityProducers.isEmpty()) {
//            addTableToDocument(document, "Производители товаров", commodityProducers, new String[]{"Наименование ССП", "Количество", "Производитель", "Статус", "Регион", "СЗТП"});
//        }
//
//        // Add "Адрес" section
//        RegAddressUlEntity regAddressUlEntity = result.getRegAddressUlEntities();
//        if (regAddressUlEntity != null) {
//            addTableToDocument(document, "Адрес", Arrays.asList(regAddressUlEntity), new String[]{"Дата регистрации", "Название организации (на русском)", "Регион регистрации (на русском)", "Район регистрации (на русском)", "Сельский район регистрации (на русском)", "Населенный пункт регистрации (на русском)", "Улица регистрации (на русском)", "Номер здания", "Номер блока", "Номер корпуса здания", "Офис (номер)", "Название ОКЭД (на русском)", "Статус ЮЛ", "Активный"});
//        }
//
//        // Add "Сведения об участниках ЮЛ" section
//        List<SvedenyaObUchastnikovUlEntity> svedenyaObUchastnikovUlEntities = result.getSvedenyaObUchastnikovUlEntities();
//        if (svedenyaObUchastnikovUlEntities != null && !svedenyaObUchastnikovUlEntities.isEmpty()) {
//            addTableToDocument(document, "Сведения об участниках ЮЛ", svedenyaObUchastnikovUlEntities, new String[]{"ФИО или наименование ЮЛ", "Идентификатор", "Дата регистрации", "Риск"});
//        }

        document.write(baos);
        baos.close();
    }

    private <T> void addTableToDocument(XWPFDocument document, String title, List<T> data, String[] columnHeaders) {
        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText(title);
        titleRun.setBold(true);
        titleRun.setFontSize(14);

        XWPFTable table = document.createTable();
        XWPFTableRow headerRow = table.getRow(0);
        for (String header : columnHeaders) {
            XWPFTableCell cell = headerRow.addNewTableCell();
            cell.setText(header);
        }
        for (T item : data) {
            XWPFTableRow row = table.createRow();
            Map<String, Object> itemMap = new HashMap<>();
            // Assuming a method to convert item to a map
            itemMap = convertItemToMap(item);
            int i = 0;
            for (String header : columnHeaders) {
                row.getCell(i).setText(itemMap.get(header).toString());
                i++;
            }
        }
    }
    private <T> Map<String, Object> convertItemToMap(T item) {
        Map<String, Object> map = new HashMap<>();
        Class<?> clazz = item.getClass();

        for (Field field : clazz.getDeclaredFields()) {
            field.setAccessible(true);
            try {
                Object value = field.get(item);
                map.put(field.getName(), value);
                System.out.println(field.getName());
                System.out.println(value);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return map;
    }
}