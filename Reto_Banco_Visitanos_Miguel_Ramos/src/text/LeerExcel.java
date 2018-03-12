package text;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeerExcel {
	
	static File src = new File ("C:\\Users\\mramos\\eclipse\\Eclipse-Oxigeno\\eclipse\\ProgramasJava\\Reto_Banco_Visitanos\\Files\\Data_File\\DataRetoBanco.xlsx");
	
	static int NumFilas;
	static String Link;
	static int Filas;
	static String TextObtenido;
	static String Direccion;
	
	public static int ContarFilas() throws IOException 
	{
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 =  wb.getSheet("Data");
		
		NumFilas = sheet1.getLastRowNum();
		
		wb.close();
		return NumFilas;
	}
	
	public static String CargarUrl(int i) throws IOException 
	{
		Filas = i;
		
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 =  wb.getSheet("Data");
		
		String CP = sheet1.getRow(Filas).getCell(0).getStringCellValue();
		String Url = sheet1.getRow(Filas).getCell(1).getStringCellValue();		
		
		if (CP.equals("N/A") == true || CP.equals("") == true || Url.equals("N/A") == true || Url.equals("") == true) {
			Url = "";
		}
		wb.close();
		return Url;
	}
	
	public static String CargarDireccion(int i) throws IOException 
	{
		Filas = i;
		
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 =  wb.getSheet("Data");
		
		String Direccion = sheet1.getRow(Filas).getCell(2).getStringCellValue();		
		
		if (Direccion.equals("N/A") == true || Direccion.equals("") == true) {
			Direccion = "";
		}
		wb.close();
		return Direccion;
	}
	
	public static void EscribirExcel(String TextObtenido, int i) {
		
		Filas = i;
		try {

		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 =  wb.getSheet("Data");
		
		sheet1.getRow(Filas).getCell(3).setCellValue(TextObtenido);
		FileOutputStream fout = new FileOutputStream(src);
		
		wb.write(fout);
		fout.flush();
		fout.close();
		}
		catch (Exception e)
		{
			System.out.println(e.getMessage());
		}
	}

}
