package Exam_Data;
import java.io.IOException;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;

import java.io.File;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;


class Student {
	int admissionNumber;
	float percentage;
	String name;
	int mathsScore;
	int chemistryScore;
	int physicsScore;
	static Logger logger = LogManager.getLogger(Student.class);

	Student(int admissionNumber,float percentage,String name,int mathsScore,int chemistryScore,int physicsScore) {

		
		this.admissionNumber = admissionNumber;
		this.percentage = percentage;
		this.name = name;
		this.mathsScore = mathsScore;
		this.chemistryScore = chemistryScore;
		this.physicsScore = physicsScore;
		

	}

	public void display(String mathsGrade,String chemistryGrade,String physicsGrade,Object mathsGradePoint, 
			Object chemistryGradePoint,Object physicsGradePointt ) {

		String grade = "Grade";
		String gradePoint = "Grade Point";
		String format = "\t{}: {}";
		logger.info("");
		
		logger.info("Admission No: {}", this.admissionNumber);
		logger.info("Percentage: {}", this.percentage);
		logger.info("Name: {}", this.name);
		logger.info("\tMark: {}", this.mathsScore);
		logger.info(format, grade, mathsGrade);
		logger.info(format, gradePoint, mathsGradePoint);
		logger.info("");
		logger.info("\tMark: {}", this.chemistryScore);
		logger.info(format, grade, chemistryGrade);
		logger.info(format, gradePoint, chemistryGradePoint);
		logger.info("Maths: ");
		logger.info("Physics: ");
		logger.info("\t Mark: {}", this.physicsScore);
		logger.info(format, grade, physicsGrade);
		logger.info(format, gradePoint, physicsGradePointt);
		logger.info("Chemistry: ");
		
		

	}

	public static void excelToList() throws IOException {

		try {

			Scanner scanner = new Scanner(System.in);
			File file = new File("C:\\Users\\shiva\\Downloads\\StudnetMarkList.xlsx");
					
			String path = file.getAbsolutePath();
			Excel_to_java excel = new Excel_to_java();
			excel.CellData(0, 1, path);

			int rowCount = excel.rowCount(path) - 1;

			
			List<Student> list = new ArrayList<Student>();

			int[] admissionNumber = new int[rowCount];
			float[] percentage = new float[rowCount];
			String[] name = new String[rowCount];
			
			int[] mathsScore = new int[rowCount];
			String[] mathsGrade = new String[rowCount];
			Object[] mathsGradePoint = new Object[rowCount];

			int[] chemistryScore = new int[rowCount];
			String[] chemistryGrade = new String[rowCount];
			Object[] chemistryGradePoint = new Object[rowCount];
			
			int[] physicsScore = new int[rowCount];
			String[] physicsGrade = new String[rowCount];
			Object[] physicsGradePoint = new Object[rowCount];


			

			float[] total = new float[rowCount];

			
			for (int i = 0; i < rowCount; i++) {

				name[i] = excel.CellData(i + 1, 1, path);
				admissionNumber[i] = excel.numCellData(i + 1, 0, path);
				physicsScore[i] = excel.numCellData(i + 1, 2, path);
				chemistryScore[i] = excel.numCellData(i + 1, 3, path);
				mathsScore[i] = excel.numCellData(i + 1, 4, path);
				total[i] = physicsScore[i] + chemistryScore[i] + (float) mathsScore[i];
				percentage[i] = (total[i] * 100) / 300;

				physicsGrade[i] = excel.gradeCalculation(physicsScore[i]);
				gradeAssigner(physicsScore, physicsGradePoint, i);

				chemistryGrade[i] = excel.gradeCalculation(chemistryScore[i]);
				gradeAssigner(chemistryScore, chemistryGradePoint, i);

				mathsGrade[i] = excel.gradeCalculation(mathsScore[i]);
				gradeAssigner(mathsScore, mathsGradePoint, i);

			}

			for (int j = 0; j < rowCount; j++) {

				Student s = new Student(admissionNumber[j],  percentage[j],name[j], mathsScore[j], chemistryScore[j],physicsScore[j]);
				list.add(s);

			}

			logger.info("Type \"Name\" to search by student's name or type \"admissionNo\" to search by student's admission number : ");
			String chooser = scanner.nextLine();

			if (chooser.equals("Name")) {

				logger.info("Type the name of the student :");
				String searchName = scanner.nextLine();

				for (int k = 0; k < list.size(); k++) {

					if (searchName.equals(name[k])) {

						list.get(k).display(physicsGrade[k], chemistryGrade[k], mathsGrade[k], physicsGradePoint[k],
								chemistryGradePoint[k], mathsGradePoint[k]);
					}

				}

			}

			if (chooser.equals("admissionNo")) {

				logger.info("Type the admission number of the student :");
				String addmissionNumber = scanner.nextLine();
				int admissionNum = Integer.parseInt(addmissionNumber);

				for (int m = 0; m < list.size(); m++) {

					if (admissionNumber[m] == admissionNum) {

						list.get(m).display(physicsGrade[m], chemistryGrade[m], mathsGrade[m], physicsGradePoint[m],
								chemistryGradePoint[m], mathsGradePoint[m]);

					}
				}
			}
			scanner.close();
		} catch (InvalidOperationException e) {

			logger.info("The path doesnt exist in the system or its not an excel file");
		}

	}

	public static void gradeAssigner(int[] Score, Object[] GradePoint, int i) {

		Excel_to_java excel = new Excel_to_java();
		if (Score[i] < 32) {

			GradePoint[i] = "C";

		} else {

			GradePoint[i] = excel.gradePointCalculation(Score[i]);

		}

	}

}
