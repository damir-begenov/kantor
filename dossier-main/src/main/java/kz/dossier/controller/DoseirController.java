package kz.dossier.controller;


import com.lowagie.text.*;

import kz.dossier.dto.AdditionalInfoDTO;
import kz.dossier.dto.AddressInfo;
import kz.dossier.dto.GeneralInfoDTO;
import kz.dossier.dto.UlAddressInfo;
import jakarta.servlet.http.HttpServletResponse;
import kz.dossier.modelsDossier.*;
import kz.dossier.repositoryDossier.EsfAll2Repo;
import kz.dossier.repositoryDossier.FlRelativesRepository;
import kz.dossier.repositoryDossier.MvAutoFlRepo;
import kz.dossier.repositoryDossier.NewPhotoRepo;
import kz.dossier.security.models.log;
import kz.dossier.security.repository.LogRepo;
import kz.dossier.service.FlRiskServiceImpl;
import kz.dossier.service.MyService;
import kz.dossier.service.RnService;
import kz.dossier.tools.DocxGenerator;
import kz.dossier.tools.PdfGenerator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.PageRequest;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;





@CrossOrigin(origins = "*", maxAge = 3000)
@RestController
@RequestMapping("/api/pandora/dossier")
public class DoseirController {
    @Autowired
    NewPhotoRepo newPhotoRepo;
    @Autowired
    EsfAll2Repo esfAll2Repo;
    @Autowired
    MvAutoFlRepo mvAutoFlRepo;
    @Autowired
    MyService myService;
    @Autowired
    FlRelativesRepository relativesRepository;
    @Autowired
    LogRepo logRepo;
    @Autowired
    FlRiskServiceImpl flRiskService;
    @Autowired
    PdfGenerator pdfGenerator;
    @Autowired
    DocxGenerator docxGenerator;
    @Autowired
    RnService rnService;

    @GetMapping("/sameAddressFl")
    public List<SearchResultModelFL> sameAddressFls(@RequestBody String iin) {
        return myService.getByAddressUsingIin(iin);
    }

    @GetMapping("/sameAddressUl")
    public List<SearchResultModelUl> sameAddressFls(@RequestBody UlAddressInfo params) {
        return myService.getByAddress(params);
    }

    // @GetMapping("/rnDetails")
    // public String getMethodName(@RequestParam String cadastral, @RequestParam String address) {
    //     rnService.getDetailedRnView(cadastral, address);
    //     return new String();
    // }

    @GetMapping("/generalInfo")
    public GeneralInfoDTO getGeneralInfo(@RequestParam String iin) {
        return myService.generalInfoByIin(iin);
    }

    @GetMapping("/additionalInfo")
    public AdditionalInfoDTO getAdditionalInfo(@RequestParam String iin) {
        return myService.additionalInfoByIin(iin);
    }


    @GetMapping("/profile")
    public NodesFL getProfile(@RequestParam String iin) {
        return myService.getNode(iin);
    }

    @GetMapping("/getRiskByIin")
    public FLRiskDto getRisk(@RequestParam String iin){
        return flRiskService.findFlRiskByIin(iin);
    }
    @GetMapping("/getFirstRowByIin")
    public FlFirstRowDto getFirstRow(@RequestParam String iin){
        return myService.getFlFirstRow(iin);
    }
    @GetMapping("/cc")
    public NodesUL getChfc(@RequestParam String bin) {
        NodesUL ss = myService.getNodeUL(bin);
        return ss;
    }
    @GetMapping("/taxpage")
    public List<TaxOutEntity> getTax(@RequestParam String bin, @RequestParam(required = false,defaultValue = "0") int page, @RequestParam(required = false,defaultValue = "10") int size) {
        return myService.taxOutEntities(bin,PageRequest.of(page,size));
    }

    @GetMapping("/pensionUl")
    public List<Map<String, Object>> pensionUl(@RequestParam String bin, @RequestParam String year, @RequestParam(required = false,defaultValue = "0") int page, @RequestParam(required = false,defaultValue = "10") int size) {
        return myService.pensionEntityUl(bin, year, PageRequest.of(page,size));
    }

    @GetMapping("/pensionsbyyear")
    public List<Map<String,Object>> pensionUl1(@RequestParam String bin, @RequestParam Double year, @RequestParam Integer page) {
        return myService.pensionEntityUl1(bin, year, page);
    }
    @GetMapping("/hierarchy")
    public FlRelativesLevelDto hierarchy(@RequestParam String iin) throws SQLException {
        return myService.createHierarchyObject(iin);
    }
    @GetMapping("/iin")
    public List<SearchResultModelFL> getByIIN(@RequestParam String iin, @RequestParam String email) throws IOException {
        List<SearchResultModelFL> fl = myService.getByIIN_photo(iin);
        log log = new log();
        log.setDate(LocalDateTime.now());
        log.setObwii("Искал в досье " + email + ": " + iin);
        log.setUsername(email);
        logRepo.save(log);
        return myService.getByIIN_photo(iin);
    }

    @GetMapping("/nomer_doc")
    public List<SearchResultModelFL> getByDoc(@RequestParam String doc) {
        return myService.getByDoc_photo(doc);
    }
    @GetMapping("/bydoc_number")
    public List<SearchResultModelFL> getByDocNumber(@RequestParam String doc_number) {
        return myService.getByDocNumber_photo(doc_number);
    }

    @GetMapping("/additionalfio")
    public List<SearchResultModelFL> getByAdditions(@RequestParam HashMap<String, String> req) {
        System.out.println(req);
        return myService.getWIthAddFields(req);
    }

    @GetMapping("/byphone")
    public List<SearchResultModelFL> getByPhone(@RequestParam String phone) {
        return myService.getByPhone(phone);
    }   @GetMapping("/byvinkuzov")
    public List<SearchResultModelFL> getByVinKuzov(@RequestParam String vin) {
        return myService.getByVinFl(vin.toUpperCase());
    }
    @GetMapping("/byvinkuzovul")
    public List<SearchResultModelUl> getByVinKuzovUl(@RequestParam String vin) {
        return myService.getByVinUl(vin.toUpperCase());
    }

    @GetMapping("/fio")
    public List<SearchResultModelFL> findByFIO(@RequestParam String i, @RequestParam String o, @RequestParam String f) {
        log log = new log();
        log.setDate(LocalDateTime.now());
//        log.setObwii("Искал в досье " + email + ": " + f + " " + i + " " + o);
//        log.setUsername(email);
        logRepo.save(log);
        return myService.getByFIO_photo(i.replace('$', '%'), o.replace('$', '%'), f.replace('$', '%'));
    }

    @GetMapping("/bin")
    public List<SearchResultModelUl> findByBin(@RequestParam String bin, @RequestParam String email) {
        log log = new log();
        log.setDate(LocalDateTime.now());
        log.setObwii("Искал в досье " + email + ": " + bin);
        log.setUsername(email);
        logRepo.save(log);
        return myService.searchResultUl(bin);
    }

    @GetMapping("/binname")
    public List<SearchResultModelUl> findBinByName(@RequestParam String name) {
        return myService.searchUlByName(name.replace('$', '%'));
    }

    @GetMapping(value = "/downloadFlPdf/{iin}", produces = MediaType.APPLICATION_PDF_VALUE)
    public byte[] generatePdfFile(HttpServletResponse response, @PathVariable("iin")String iin)throws IOException, DocumentException {
        response.setContentType("application/pdf");
        String headerkey = "Content-Disposition";
        String headervalue = "attachment; filename=doc" + ".pdf";
        response.setHeader(headerkey, headervalue);
        NodesFL r =  myService.getNode(iin);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        pdfGenerator.generate(r, baos);
        return baos.toByteArray();
    }

    @GetMapping("/downloadFlDoc/{iin}")
    public byte[] generateDoc(@PathVariable String iin, HttpServletResponse response) throws IOException, InvalidFormatException {
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        String headerkey = "Content-Disposition";
        String headervalue = "attachment; filename=document.docx";
        response.setHeader(headerkey,headervalue);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        NodesFL result =  myService.getNode(iin);
        docxGenerator.generateDoc(result,baos);
        return baos.toByteArray();
    }

    @GetMapping(value = "/downloadUlPdf/{bin}", produces = MediaType.APPLICATION_PDF_VALUE)
    public byte[] generateUlPdfFile(HttpServletResponse response, @PathVariable("bin")String bin) throws IOException, DocumentException {
        response.setContentType("application/pdf");
        String headerkey = "Content-Disposition";
        String headervalue = "attachment; filename=doc" + ".pdf";
        response.setHeader(headerkey, headervalue);
        NodesUL r =  myService.getNodeUL(bin);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        pdfGenerator.generate(r, baos);
        return baos.toByteArray();
    }
    @GetMapping(value = "/downloadUlDoc/{bin}")
    public byte[] generateUlWordFile(HttpServletResponse response, @PathVariable("bin")String bin) throws IOException, DocumentException {
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        String headerkey = "Content-Disposition";
        String headervalue = "attachment; filename=document.docx";
        response.setHeader(headerkey,headervalue);
        NodesUL r =  myService.getNodeUL(bin);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        docxGenerator.generateUl(r, baos);
        return baos.toByteArray();
    }
}
