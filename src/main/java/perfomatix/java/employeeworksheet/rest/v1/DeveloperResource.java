package perfomatix.java.employeeworksheet.rest.v1;

import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import perfomatix.java.employeeworksheet.dto.DeveloperDetailsDTO;
import perfomatix.java.employeeworksheet.dto.ResponseDTO;
import perfomatix.java.employeeworksheet.service.DownloadReportService;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@RequestMapping("api/v1")
@RequiredArgsConstructor
@RestController
public class DeveloperResource {
    private final DownloadReportService downloadReportService;

    @PostMapping(value = " /generate/excel/worksheet")
    public ResponseEntity<?> generateExcelSheet(@RequestBody List<DeveloperDetailsDTO> developerDetailsList) {
        ByteArrayOutputStream bytes = downloadReportService.generateExcelSheet(developerDetailsList);
        String fileName = "Developer Details Excel.xlsx";
        return ResponseEntity.ok().contentType(MediaType.parseMediaType(MediaType.APPLICATION_OCTET_STREAM_VALUE)).
                header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"").body(bytes.toByteArray());
    }

}
