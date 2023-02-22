package perfomatix.java.employeeworksheet.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.springframework.stereotype.Service;
import perfomatix.java.employeeworksheet.dto.DeveloperDetailsDTO;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.UUID;
@Slf4j
@Service
public class DownloadReportService {
    private static String[] columns = {"DeveloperName", "TaskName", "Taskstaus", "WorkedTime", "EstimatedEffort"};

    public  ByteArrayOutputStream generateExcelSheet(List<DeveloperDetailsDTO> developerDetailsList)  {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {

            SXSSFWorkbook workbook = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
            SXSSFSheet sheet = workbook.createSheet("developerDetailsList");
            CreationHelper creationHelper = workbook.getCreationHelper();

            // set font
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            //hyperlink
            CellStyle hyperLinkStyle = workbook.createCellStyle();
            Font hyperLinkFont = workbook.createFont();
            hyperLinkFont.setUnderline(XSSFFont.U_SINGLE);
            hyperLinkStyle.setFont(hyperLinkFont);

            // header columns
            Row headerRow = sheet.createRow(0);

            //data cell style
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setWrapText(true);

            //track column for sizing
            sheet.trackAllColumnsForAutoSizing();

            //auto size the header column
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerCellStyle);
                sheet.autoSizeColumn(i);
            }

            int rowNum = 1;

            for (int i = developerDetailsList.size() - 1; i >= 0; i--) {
                try {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(developerDetailsList.get(i).developerName());
                    row.createCell(1).setCellValue(developerDetailsList.get(i).taskName());
                    row.createCell(2).setCellValue(developerDetailsList.get(i).taskStatus());
                    row.createCell(3).setCellValue(developerDetailsList.get(i).workedTime());
                    row.createCell(4).setCellValue(developerDetailsList.get(i).estimateEffort());

                } catch (Exception e) {
                    log.error("Error occurred inside loop while writing rows " + developerDetailsList.get(i).taskName(), e);

                }
            }
            //auto sizing the data cells
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerCellStyle);
                sheet.autoSizeColumn(i);
            }

            //removing cell tracking
            sheet.untrackAllColumnsForAutoSizing();

            workbook.write(outputStream);
            workbook.close();



        } catch (Exception e) {
        log.warn("Failed to generate excel sheet", e);
        }
        return outputStream;
}
}

