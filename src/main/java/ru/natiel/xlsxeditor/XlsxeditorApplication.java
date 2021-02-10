package ru.natiel.xlsxeditor;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import java.io.IOException;

@SpringBootApplication
public class XlsxeditorApplication {

	public static void main(String[] args) throws IOException {
		ConfigurableApplicationContext context  = SpringApplication.run(XlsxeditorApplication.class, args);
		StartupService service = context.getBean(StartupService.class);
		service.start();
		context.close();
	}

}
