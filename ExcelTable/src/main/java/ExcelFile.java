import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.File;
import java.io.FileOutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import org.json.JSONException;
import org.json.JSONObject;
import java.nio.charset.StandardCharsets;
import java.util.GregorianCalendar;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import static java.util.Calendar.*;


class JsonParser {
	
	String convertStreamToString(InputStream input)throws IOException{
	BufferedReader reader = new BufferedReader(new InputStreamReader(input, StandardCharsets.UTF_8));
        StringBuilder string = new StringBuilder();

        String line;
        while ((line = reader.readLine()) != null) {
            string.append(line).append("\n");
        }
        input.close();

	return string.toString();}
	
	private String GetRegion(String index) {
		try{String request = "https://www.postindexapi.ru/json/"+index.substring(0,3)+"/"+index+".json";
		URL url = new URL(request);
		HttpURLConnection connection = (HttpURLConnection)url.openConnection();
		connection.setRequestProperty("Content-Type", "application/json;charset=windows-1251");
		connection.connect();
		InputStream input = connection.getInputStream();
		String response = convertStreamToString(input);
		JSONObject userJson = new JSONObject(response);
		return userJson.getString("Region");
		}catch(Exception exception){return "-";}
		}
		

    Person GetPerson(String response)throws JSONException {
        JSONObject userJson = new JSONObject(response);
        String surname = userJson.getString("lname");
        String name = userJson.getString("fname");
        String patrynomic = userJson.getString("patronymic");
        String gender = userJson.getString("gender");
		if (gender.equals("w"))
			{gender = "ЖЕН";}
		else
			{gender = "МУЖ";}
        String postcode = userJson.getString("postcode");
        String city = userJson.getString("city");
        String street = userJson.getString("street");
		int house = userJson.getInt("house");
		int flat = userJson.getInt("apartment");
		String region = GetRegion(postcode);
		return new Person(surname,name,patrynomic,gender,city,postcode,region,city,street,house,flat);

    }
}
class HttpRequest{
	private final JsonParser jsonParser;
	
	HttpRequest(){
	jsonParser = new JsonParser();
	}
	
	Person GetPersonInformation()throws IOException,JSONException{
	String request = "https://randus.org/api.php";
	URL url = new URL(request);
	HttpURLConnection connection = (HttpURLConnection)url.openConnection();
	connection.setRequestProperty("Content-Type", "application/json;charset=windows-1251");
	connection.connect();
	InputStream input = connection.getInputStream();
	String response = jsonParser.convertStreamToString(input);
		return jsonParser.GetPerson(response);
	}	
}
class Person{
	private String name;
	private String surname;
	private String patronymic;
	private String postcode;
	private String region;
	private String city;
	private String street;
	private int house;
	private int flat;
	private String personGender;
	private int personAge;
	private Birthday personBirthday;
	private String personBirthPlace;

		Person(String surname,String name,String patronymic,String gender,
		String birthPlace,String postcode,String region,
		String city,String street,int house,int flat)
		{
		this.surname = surname;
		this.name = name;
		this.patronymic = patronymic;
		this.personGender = gender;
		this.personBirthPlace = birthPlace;
		this.personAge = (int)Math.round(Math.random()*100);
		this.personBirthday = new Birthday(this.personAge);
		this.postcode = postcode;
			this.region = region;
		this.city = city;
		this.street = street;
		this.house = house;
		this.flat = flat;
		}
	
	String GetName(){
	return this.name;}
	
	String GetSurname(){
	return this.surname;}
	
	String GetPatronymic(){
	return this.patronymic;}
	
	String GetGender(){
	return this.personGender;}
	
	int GetAge(){
	return this.personAge;}
	
	String GetBirthday(){
	return this.personBirthday.GetData();}
	
	String GetBirthPlace(){
	return this.personBirthPlace;}
	
	String GetPostcode(){
	return this.postcode;}
	
	String GetRegion(){
	return this.region;}
	
	String GetCity(){
	return this.city;}
	
	String GetStreet(){
	return this.street;}
		
	int GetHouse(){
	return this.house;}
		
	int GetFlat(){
	return this.flat;}
	
	Person(Person person){
		

		this.surname = person.surname;
		this.name = person.name;
		this.patronymic = person.patronymic;
		this.personGender = person.personGender;
		this.personBirthPlace = person.personBirthPlace;
		this.personBirthday = person.personBirthday;
		this.personAge = person.personAge;
		this.postcode = person.postcode;
		this.region = person.region;
		this.city = person.city;
		this.street = person.street;
		this.house = person.house;
		this.flat = person.flat;
	}
	
}


class Birthday{
    private GregorianCalendar birthday;

    Birthday(int age)
        
    {
        this.birthday = new GregorianCalendar();
        this.birthday.set(YEAR,2019 - age);
        int day = randBetween(this.birthday.getActualMaximum(DAY_OF_YEAR));
        this.birthday.set(DAY_OF_YEAR,day);
    }

	private static int randBetween(int end){
	return 1 + (int)Math.round(Math.random()*(end- 1));}
	
	String GetData(){
		String data = "";
		if (this.birthday.get(DAY_OF_MONTH)<10){
			data+="0";
		}
		data+=this.birthday.get(DAY_OF_MONTH)+"-";
		if(this.birthday.get(MONTH)<10){
			data+="0";
		}
		data+=this.birthday.get(MONTH)+"-";
		data+=this.birthday.get(YEAR);
		return data;
		}
		
}

	public class ExcelFile {
    public static void main(String[] args) throws JSONException {
		try {
			String[] data;
			data = new String[]{"Имя","Фамилия","Отчество","Возраст",
					"Пол","Дата рождения","Место рождения",
					"Индекс","Страна","Область","Город","Улица",
					"Дом","Квартира"};
			Scanner sc = new Scanner(System.in);
			System.out.println("Введите число:");
			int number = sc.nextInt();
			sc.close();
			int randomFileNumber = (int) Math.round(Math.random() * (100));
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Лист");
			int rowNumber = 0;
			Row row = sheet.createRow(rowNumber);
			for(int i = 0; i < data.length; i++) {
				row.createCell(i).setCellValue(data[i]);
			}
			for (int i = 0; i < number; i++) {
				HttpRequest request = new HttpRequest();
				Person person = new Person(request.GetPersonInformation());
				createSheetHeader(sheet, ++rowNumber, person);
			}
			File file = new File(System.getProperty("user.dir") + "\\ExcelRandomHumans" + randomFileNumber + ".xls");
			boolean fileCreated = file.getParentFile().mkdirs();
			if (!fileCreated) {
				FileOutputStream outFile;
				outFile = new FileOutputStream(file);
				workbook.write(outFile);
				System.out.println("Excel файл успешно создан.Путь: " + file.getAbsolutePath());
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

	}



    private static void createSheetHeader(HSSFSheet sheet, int rowNum, Person person) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(person.GetName());
        row.createCell(1).setCellValue(person.GetSurname());
        row.createCell(2).setCellValue(person.GetPatronymic());
        row.createCell(3).setCellValue(person.GetAge());
		row.createCell(4).setCellValue(person.GetGender());
        row.createCell(5).setCellValue(person.GetBirthday());
        row.createCell(6).setCellValue(person.GetBirthPlace());
        row.createCell(7).setCellValue(person.GetPostcode());
		row.createCell(8).setCellValue("Россия");
        row.createCell(9).setCellValue(person.GetRegion());
        row.createCell(10).setCellValue(person.GetCity());
        row.createCell(11).setCellValue(person.GetStreet());
		row.createCell(12).setCellValue(person.GetHouse());
        row.createCell(13).setCellValue(person.GetFlat());

    }
 
}
 