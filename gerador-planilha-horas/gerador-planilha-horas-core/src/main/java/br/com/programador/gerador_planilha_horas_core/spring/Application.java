package br.com.programador.gerador_planilha_horas_core.spring;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.support.AbstractApplicationContext;

import br.com.programador.gerador_planilha_horas_core.excel.ExcelUtil;
import br.com.programador.gerador_planilha_horas_core.rn.ExcelRN;

@Configuration
@ComponentScan("br.com.programador.gerador_planilha_horas_core")
public class Application {

	@Bean
	public ExcelRN excelRN() {
		return new ExcelRN();
	}

	@Bean
	public ExcelUtil excelUtil() {
		return new ExcelUtil();
	}

	public static void main(String[] args) {
		AbstractApplicationContext context = new AnnotationConfigApplicationContext(Application.class);
		ExcelRN excelRN = context.getBean(ExcelRN.class);
		try {
			excelRN.abrirDocumento();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
		context.close();
	}

}
