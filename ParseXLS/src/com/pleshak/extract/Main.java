package com.pleshak.extract;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.pleshak.extract.domain.Person;

public class Main {

	public static void main(String[] args) throws Exception {

		// //////////////////
		final String fileName = "file.xls";
		ArrayList<Person> list = (ArrayList<Person>) readExelFileData(fileName);
		Iterator<Person> itr = list.iterator();
		while (itr.hasNext()) {
			Person person = itr.next();

			System.out.print(person.getName() + "  ");
			System.out.println(person.getAge());
		}

	}

	private static List<Person> readExelFileData(String fileName) throws Exception {

		List<Person> personList = new ArrayList<Person>();

		InputStream in = new FileInputStream(fileName);

		Workbook workbook = null;

		if (fileName.toLowerCase().endsWith("xlsx")) {

			workbook = new XSSFWorkbook(in);
		} else if (fileName.toLowerCase().endsWith("xls")) {

			workbook = new HSSFWorkbook(in);
		}

		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {
			Person aperson = new Person();
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.iterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				int cellType = cell.getColumnIndex();

				if (cellType == 0) {
					// System.out.print(cell.getStringCellValue());
					aperson.setName((String) cell.getStringCellValue());
				}

				else if (cellType == 1) {
					// System.out.println(cell.getNumericCellValue());
					aperson.setAge(cell.getNumericCellValue());
				}

			}
			personList.add(aperson);
		}

		// System.out.println();

		return personList;
	}
}
