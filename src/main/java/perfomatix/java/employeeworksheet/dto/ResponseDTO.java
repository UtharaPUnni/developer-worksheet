package perfomatix.java.employeeworksheet.dto;

import org.springframework.http.HttpStatus;

public record ResponseDTO<T>(HttpStatus code, String status, String message, T data){}
