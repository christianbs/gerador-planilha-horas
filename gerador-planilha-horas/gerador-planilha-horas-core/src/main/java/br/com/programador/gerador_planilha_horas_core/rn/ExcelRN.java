package br.com.programador.gerador_planilha_horas_core.rn;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;

import br.com.programador.gerador_planilha_horas_core.excel.ExcelUtil;

public class ExcelRN {

	@Autowired
	private ExcelUtil util;

	public void teste() {
		util.teste();
	}

	public void abrirDocumento() throws IOException, EncryptedDocumentException, InvalidFormatException {
		util.gerarPlanilha();
	}

}
