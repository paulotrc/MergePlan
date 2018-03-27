package br.com.paulotrc.mergeplan;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import javax.swing.JOptionPane;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExecutarPlan {

	private InputStream excelFileToRead = null;
	private FileOutputStream excelFileToWrite = null;
	private String nomeArquivo;
	private String[] arrayColunaZERO;
	private String[][] arrayLinhasZERO;
	
	private String[] arrayColunaUM;
	private String[][] arrayLinhasUM;

	public void ExecutarPlanilha(File arquivo) throws IOException{
		excelFileToRead = new FileInputStream(arquivo.getAbsolutePath());
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHHmmss");
		nomeArquivo = arquivo.getAbsolutePath().toLowerCase().replace(".xlsx", "_"+df.format(new Date())) +".xlsx";
		excelFileToWrite = new FileOutputStream(new File(nomeArquivo));
	}

	@SuppressWarnings("rawtypes")
	public void readXLSFile() throws Exception
	{
		JOptionPane.showMessageDialog(null,"Iniciando processo de leitura.");
		XSSFWorkbook  wb = new XSSFWorkbook(excelFileToRead);
		excelFileToRead = null;
		
		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row = null; 
		XSSFCell cell = null;

		Iterator rows = sheet.rowIterator();
		int tempCol = 0;
		int tempCel = 0;
		arrayLinhasZERO = new String[sheet.getLastRowNum()+1][sheet.getLastRowNum()+1];
		arrayColunaZERO = new String[sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()];
		row = null;
		Iterator cells = null;
		int contagemLinhas = 0;
		while (rows.hasNext())
		{
			row=(XSSFRow) rows.next();
			cells = row.cellIterator();
			while (cells.hasNext())
			{
				cell=(XSSFCell) cells.next();
		
				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
				{
					arrayColunaZERO[tempCol] = cell.getStringCellValue();
					tempCol++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
				{
					arrayColunaZERO[tempCol] = new Integer(new Double(cell.getNumericCellValue()).intValue()).toString();
					tempCol++;
				}
			}
			arrayLinhasZERO[tempCel] = arrayColunaZERO;
			tempCel++;
			arrayColunaZERO = new String[sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()];
			tempCol = 0;
			int valLinhas = contagemLinhas++;
			int val = ((valLinhas) * 100)/(sheet.getLastRowNum()+1);
		}
		
		arrayColunaZERO = null;
		
		sheet = null;
		sheet = wb.getSheetAt(1);
		rows = null;
		rows = sheet.rowIterator();
		tempCol = 0;
		tempCel = 0;
		arrayLinhasUM = new String[sheet.getLastRowNum()+1][sheet.getLastRowNum()+1];
		arrayColunaUM = new String[sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()];
		row = null;
		cells = null;

		contagemLinhas = 0;
		while (rows.hasNext())
		{
			row=(XSSFRow) rows.next();
			cells = row.cellIterator();
			while (cells.hasNext())
			{
				cell=(XSSFCell) cells.next();
		
				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
				{
					arrayColunaUM[tempCol] = cell.getStringCellValue();
					tempCol++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
				{
					arrayColunaUM[tempCol] = new Integer(new Double(cell.getNumericCellValue()).intValue()).toString();
					tempCol++;
				}
				else
				{
				}
			}
			arrayLinhasUM[tempCel] = arrayColunaUM;
			tempCel++;
			arrayColunaUM = new String[sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()];
			tempCol = 0;
			
			int val = ((contagemLinhas++) * 100)/(sheet.getLastRowNum()+1);
		}
		
		wb = null;
		sheet = null;
		
		JOptionPane.showMessageDialog(null,"Registros lidos.");
	}
	
	public void writeXLSFile() throws Exception {
		JOptionPane.showMessageDialog(null,"Iniciando processo de escrita.");
		String sheetName = "Sheet1";//name of sheet

		SXSSFWorkbook wb = new SXSSFWorkbook();
		SXSSFSheet sheet = (SXSSFSheet) wb.createSheet(sheetName) ;

		int contagemLinhas = 0;
		int valorIteracao = 0;
		for (int a = 0; a < arrayLinhasZERO.length; a++) {
			SXSSFRow row = (SXSSFRow) sheet.createRow(sheet.getLastRowNum());
			while (valorIteracao < arrayLinhasUM.length) {
				int lastCell = 0;
				obterCelulasDaLinha(row, arrayLinhasZERO[a]);
				lastCell = row.getLastCellNum();
				obterCelulasDaLinha(row, arrayLinhasUM[valorIteracao], lastCell);
				valorIteracao++;
				row = (SXSSFRow) sheet.createRow(sheet.getLastRowNum()+1);
			}
			valorIteracao = 0;
//			System.out.println(a+1);
			int val = ((contagemLinhas++) * 100)/(sheet.getLastRowNum()+1);
		}

		JOptionPane.showMessageDialog(null,"Escrevendo no arquivo.");
		wb.write(excelFileToWrite);
		JOptionPane.showMessageDialog(null,"Escrita finalizada.");
		JOptionPane.showMessageDialog(null,"Finalizando arquivo.");
		excelFileToWrite.flush();
		excelFileToWrite.close();
		JOptionPane.showMessageDialog(null,"Arquivo finalizado disponivel em: ["+nomeArquivo+"].");
		
	}

	private void obterCelulasDaLinha(SXSSFRow row, String[] arrayCelulas) {
		for (int i = 0; i < arrayCelulas.length; i++) {
			SXSSFCell cell = (SXSSFCell) row.createCell(i);
			cell.setCellValue(arrayCelulas[i]);
		}
		
	}

	private void obterCelulasDaLinha(SXSSFRow row, String[] arrayCelulas, int lastCell) {
		for (int i = 0; i < arrayCelulas.length; i++) {
			SXSSFCell cell = (SXSSFCell) row.createCell(lastCell);
			cell.setCellValue(arrayCelulas[i]);
			lastCell++;
		}
		
	}
}
