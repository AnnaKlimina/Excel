import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.ParseException;
import java.util.GregorianCalendar;
import java.util.Scanner;




class JsonParser {

    public String convertStreamToString(InputStream input)throws IOException{
        BufferedReader reader = new BufferedReader(new InputStreamReader(input,"utf-8"));
        StringBuilder string = new StringBuilder();

        String line;
        while ((line = reader.readLine()) != null) {
            string.append(line).append("\n");
        }
        input.close();

        return string.toString();}

    public String GetRegion(String index) throws IOException,JSONException{
        String request = "https://www.postindexapi.ru/json/"+index.substring(0,3)+"/"+index+".json";
        URL url = new URL(request);
        HttpURLConnection connection = (HttpURLConnection)url.openConnection();
        connection.setRequestProperty("Content-Type", "application/json;charset=windows-1251");
        connection.connect();
        InputStream input = connection.getInputStream();
        String response = convertStreamToString(input);
        JSONObject userJson = new JSONObject(response);
        String region = userJson.getString("Region");
        return region;
    }


    public Person GetPerson(String response)throws IOException,JSONException {
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
    final JsonParser jsonParser;

    public HttpRequest(){
        jsonParser = new JsonParser();
    }

    public Person GetPersonInformation()throws IOException,JSONException{
        String request = "https://randus.org/api.php";
        URL url = new URL(request);
        HttpURLConnection connection = (HttpURLConnection)url.openConnection();
        connection.setRequestProperty("Content-Type", "application/json;charset=windows-1251");
        connection.connect();
        InputStream input = connection.getInputStream();
        String response = jsonParser.convertStreamToString(input);
        Person person = jsonParser.GetPerson(response);
        return person;
    }
}
class Person{
    String name;
    String surname;
    String patronymic;
    String postcode;
    String country;
    String region;
    String city;
    String street;
    int house;
    int flat;
    String personGender;
    int personAge;
    Birthday personBirthday;
    String personBirthPlace;

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
        this.country = "Россия";
        this.region = region;
        this.city = city;
        this.street = street;
        this.house = house;
        this.flat = flat;
    }

    public String GetName(){
        return this.name;}

    public String GetSurname(){
        return this.surname;}

    public String GetPatronymic(){
        return this.patronymic;}

    public String GetGender(){
        return this.personGender;}

    public int GetAge(){
        return this.personAge;}

    public String GetBirthday(){
        return this.personBirthday.GetData();}

    public String GetBirthPlace(){
        return this.personBirthPlace;}

    public String GetPostcode(){
        return this.postcode;}

    public String GetRegion(){
        return this.region;}

    public String GetCity(){
        return this.city;}

    public String GetStreet(){
        return this.street;}

    public int GetHouse(){
        return this.house;}

    public int GetFlat(){
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
        this.country = "Россия";
        this.region = person.region;
        this.city = person.city;
        this.street = person.street;
        this.house = person.house;
        this.flat = person.flat;
    }

}


class Birthday{
    GregorianCalendar birthday;

    Birthday(int age)

    {
        this.birthday = new GregorianCalendar();
        this.birthday.set(this.birthday.YEAR,2019 - age);
        int day = randBetween(1, this.birthday.getActualMaximum(this.birthday.DAY_OF_YEAR));
        this.birthday.set(this.birthday.DAY_OF_YEAR,day);
    }

    public static int randBetween(int start, int end){
        return start + (int)Math.round(Math.random()*(end-start));}

    public String GetData(){
        String data = "";
        if (this.birthday.get(this.birthday.DAY_OF_MONTH)<10){
            data+="0";
        }
        data+=this.birthday.get(this.birthday.DAY_OF_MONTH)+"-";
        if(this.birthday.get(this.birthday.MONTH)<10){
            data+="0";
        }
        data+=this.birthday.get(this.birthday.MONTH)+"-";
        data+=this.birthday.get(this.birthday.YEAR);
        return data;
    }

}

public class ExcelRandomHumans {
    public static void main(String[] args) throws ParseException,JSONException {
        try{
            Scanner sc = new Scanner(System.in);
            System.out.println("Введите номер:");
            int number = sc.nextInt();
            sc.close();
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Лист");
            int rowNumber = 0;
            Row row = sheet.createRow(rowNumber);
            row.createCell(0).setCellValue("Имя");
            row.createCell(1).setCellValue("Фамилия");
            row.createCell(2).setCellValue("Отчество");
            row.createCell(3).setCellValue("Возраст");
            row.createCell(4).setCellValue("Пол");
            row.createCell(5).setCellValue("Дата Рождения");
            row.createCell(6).setCellValue("Место Рождения");
            row.createCell(7).setCellValue("Индекс");
            row.createCell(8).setCellValue("Страна");
            row.createCell(9).setCellValue("Область");
            row.createCell(10).setCellValue("Город");
            row.createCell(11).setCellValue("Улица");
            row.createCell(12).setCellValue("Дом");
            row.createCell(13).setCellValue("Квартира");

            for (int i = 0; i < number; i++) {
                HttpRequest request = new HttpRequest();
                Person person = new Person(request.GetPersonInformation());
                createSheetHeader(sheet, ++rowNumber, person);
            }


            int randomFileNumber = (int)Math.round(Math.random()*(100));
            File file = new File(System.getProperty("user.dir").toString()+"\\ExcelRandomHumans"+randomFileNumber+".xls");
            file.getParentFile().mkdirs();
            FileOutputStream outFile = new FileOutputStream(file);

            workbook.write(outFile);
            System.out.println("Excel файл успешно создан. Полный путь: "+file.getAbsolutePath());
        }catch(IOException e){e.printStackTrace();}
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
