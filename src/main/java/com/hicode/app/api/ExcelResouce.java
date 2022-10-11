package com.hicode.app.api;

import com.hicode.app.dto.ProductDTO;
import com.hicode.app.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequiredArgsConstructor
public class ExcelResouce {
    private final ExcelService excelService;
    @PostMapping("/upload")
    public ResponseEntity<?> upload(@RequestParam("file")MultipartFile file){
        try {
            List<ProductDTO> productDTOS = excelService.readFile(file);
            return ResponseEntity.ok(productDTOS);
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(e.getMessage());
        }
    }
    @GetMapping("/file/{filename:.+}")
    public ResponseEntity<Resource> download(@PathVariable("filename")String filename){
        try {
            InputStreamResource file = new InputStreamResource(excelService.downloadFile(filename));
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(file);
        }
         catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

}
