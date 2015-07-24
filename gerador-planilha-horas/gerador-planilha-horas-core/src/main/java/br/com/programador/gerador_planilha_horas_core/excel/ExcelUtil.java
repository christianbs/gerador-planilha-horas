package br.com.programador.gerador_planilha_horas_core.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import br.com.programador.gerador_planilha_horas_core.constante.Horario;

public class ExcelUtil {
	private int numeroGerado;
	private String caminhoArquivoGerado;
	private String caminhoArquivoOriginal;
	private Sheet folha;
	private Workbook wb;
	private FileInputStream arquivo;

	public ExcelUtil() {
		caminhoArquivoOriginal = "src/main/resources/FOLHA DE PONTO - 7COMM.xlsx";
		caminhoArquivoGerado = "C:/Users/cristian.sticchi/Desktop/FOLHA DE PONTO - Agosto 2015 - 7COMM.xlsx";
	}

	public void teste() {
		System.out.println("Hello");
	}

	private void abrirDocumento() throws IOException, EncryptedDocumentException, InvalidFormatException {
		arquivo = new FileInputStream(new File(caminhoArquivoOriginal));
		wb = WorkbookFactory.create(arquivo);
		folha = wb.getSheetAt(0);
	}

	private void fecharDocumento() throws IOException {
		arquivo.close();
	}

	public void gerarPlanilha() throws IOException, EncryptedDocumentException, InvalidFormatException {
		abrirDocumento();
		for (int contador = 11; contador < 42; contador++) {
			preencherLinha(folha.getRow(contador));
		}
		fecharDocumento();
		gravarAlteracoes();
	}

	private void gravarAlteracoes() throws IOException {
		FileOutputStream outFile = new FileOutputStream(new File(caminhoArquivoGerado));
		wb.write(outFile);
		outFile.close();
	}

	private void preencherLinha(Row linha) {
		Cell celula = linha.getCell(2);
		if (!celula.getStringCellValue().equals("SÁB") && !celula.getStringCellValue().equals("DOM")) {

			gerarNumero();
			celula = linha.getCell(4);
			celula.setCellValue(formatarHorario(Horario.ENTRADA.getHorario(), numeroGerado));
			celula = linha.getCell(8);
			celula.setCellValue(formatarHorario(Horario.SAIDA.getHorario(), numeroGerado));

			gerarNumero();
			celula = linha.getCell(5);
			celula.setCellValue(formatarHorario(Horario.SAIDA_ALMOCO.getHorario(), numeroGerado));
			celula = linha.getCell(7);
			celula.setCellValue(formatarHorario(Horario.RETORNO_ALMOCO.getHorario(), numeroGerado));

			celula = linha.getCell(13);
			celula.setCellValue(Horario.ENTRADA.getHorario());
		}
	}

	private String formatarHorario(String horarioCompleto, int minutos) {
		String minutosEmTexto = String.valueOf(minutos);
		if (minutosEmTexto.length() < 2) {
			minutosEmTexto = "0" + minutosEmTexto;
		}
		horarioCompleto = horarioCompleto.replaceFirst("00", minutosEmTexto);
		return horarioCompleto;
	}

	private void gerarNumero() {
		numeroGerado = new Random().nextInt(30);
	}

}
