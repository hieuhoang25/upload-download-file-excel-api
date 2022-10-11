package com.hicode.app.service;

import com.hicode.app.dto.ProductDTO;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface ExcelService {
    List<ProductDTO> readFile(MultipartFile file) throws IOException;
    void saveDatatoDB();
    ByteArrayInputStream  downloadFile(String filename) throws IOException;


}
