package com.sobngwi.springbootexcel.rest;


import com.sobngwi.springbootexcel.model.User;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;

import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

@RestController
public class Export {


    @Autowired
   // UserService userService;

    /**
     * Handle request to download an Excel document
     */
    @RequestMapping(value = "/downloads", method = RequestMethod.GET)
    public  String download(Model model) {
        model.addAttribute("users", Arrays.asList(new User())/*userService.findAllUsers()*/);
        return "OK";
    }

    @RequestMapping(value="/download", method=RequestMethod.GET)
    public void getDownload(HttpServletResponse response) throws Exception{

        HSSFWorkbook wb = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("monfichier.xls");
        final File file1 = new File("toto");

        // Lecture fichier original
        //final Workbook workbook = WorkbookFactory.create(fileOut.);
        // create excel xls sheet
        Sheet sheet = wb.createSheet("User Detail");
        sheet.setDefaultColumnWidth(30);

        // create style for header cells
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Arial");
        style.setFillForegroundColor(HSSFColor.BLUE.index);
        style.setFillPattern(SOLID_FOREGROUND);
        font.setBold(true);
        font.setColor(HSSFColor.WHITE.index);
        style.setFont(font);


        // create header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Firstname");
        header.getCell(0).setCellStyle(style);
        header.createCell(1).setCellValue("LastName");
        header.getCell(1).setCellStyle(style);
        header.createCell(2).setCellValue("Age");
        header.getCell(2).setCellStyle(style);
        header.createCell(3).setCellValue("Job Title");
        header.getCell(3).setCellStyle(style);
        header.createCell(4).setCellValue("Company");
        header.getCell(4).setCellStyle(style);
        header.createCell(5).setCellValue("Address");
        header.getCell(5).setCellStyle(style);
        header.createCell(6).setCellValue("City");
        header.getCell(6).setCellStyle(style);
        header.createCell(7).setCellValue("Country");
        header.getCell(7).setCellStyle(style);
        header.createCell(8).setCellValue("Phone Number");
        header.getCell(8).setCellStyle(style);



        int rowCount = 1;

        for(User user : Arrays.asList(new User())){
            Row userRow =  sheet.createRow(rowCount++);
            userRow.createCell(0).setCellValue(user.getFirstName());
            userRow.createCell(1).setCellValue(user.getLastName());
            userRow.createCell(2).setCellValue(user.getAge());
            userRow.createCell(3).setCellValue(user.getJobTitle());
            userRow.createCell(4).setCellValue(user.getCompany());
            userRow.createCell(5).setCellValue(user.getAddress());
            userRow.createCell(6).setCellValue(user.getCity());
            userRow.createCell(7).setCellValue(user.getCountry());
            userRow.createCell(8).setCellValue(user.getPhoneNumber());

        }

        // Get your file stream from wherever.
        //InputStream myStream = someClass.returnFile();
        //InputStream myStream = new FileInputStream(file1);
        wb.write(fileOut);
        fileOut.close();
        // Set the content type and attachment header.
        response.addHeader("Content-disposition", "attachment;filename=myfilename.txt");
        response.setContentType("application/vnd.ms-excel");

        // Copy the stream to the response's output stream.
        IOUtils.copy(new FileInputStream("monfichier.xls"), response.getOutputStream());
        response.flushBuffer();
    }

}
