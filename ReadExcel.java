import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream.GetField;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Queue;
import java.util.Scanner;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		int contadorMenu = 0;
		List<String> menu = Stream.of("lm", "name", "free_shipping", "description", "price")
				.collect(Collectors.toList());

		File excelFile = new File("javaleroy.xlsx");

		FileInputStream fis = new FileInputStream(excelFile);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		Iterator<Row> rowIt = sheet.iterator();

		Queue<JSONObject> queueCell = new LinkedList<>();
		JSONObject obj = new JSONObject();
		JSONArray allDataArray = new JSONArray();
		while (rowIt.hasNext()) {
			Row row = rowIt.next();
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {

				Cell cell = cellIterator.next();
				if (contadorMenu < 7) {
					contadorMenu++;
				} else {
					obj.put(menu.get((contadorMenu + 3) % 5), cell);
					contadorMenu++;
				}
			}

			if (contadorMenu > 7) {
				queueCell.add(obj);

				allDataArray.put(obj);
				obj = new JSONObject();
			}
		}
		
		fis.close();
		for(int i=0;i<allDataArray.length();i++) {
			System.out.println(""+i+allDataArray.get(i));
		}
		
		
		Scanner scanner = new Scanner(System.in);  
		//Para jogar em um inteiro
		System.out.println("Digite uma opção a ser removida e tecle enter");
		int escolha = scanner.nextInt();
		scanner.close();

			//ATUALIZAR, criar planilha de teste							
			// O numero da row devera ser = ou maior que 3
			Row row = sheet.getRow(escolha+3);
			sheet.removeRow(row);

	            // open an OutputStream to save written data into XLSX file
	            FileOutputStream os = new FileOutputStream("javaleroyteste.xlsx");
	            workbook.write(os);
	            System.out.println("Planilha de teste criada");
	            workbook.close();
	}	
}
