package com.pleshak.extract;

import java.util.Iterator;
import java.util.List;

import com.pleshak.extract.domain.Person;

public class Main {

	public static void main(String[] args) throws Exception {

		String fileName = "file.xls";
		ExcelReader excelReader = new ExcelReader();
		List<Person> personList = excelReader.readExelFileData(fileName);
		System.out.println(personList);
		Iterator<Person> itr = personList.iterator();
		while (itr.hasNext()) {
			Person person = itr.next();

			System.out.print(person.getName() + "  ");
			System.out.println(person.getAge());
		}

	}

}
