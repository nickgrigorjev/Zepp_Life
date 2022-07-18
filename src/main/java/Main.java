import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.query.Query;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.*;
import java.sql.Date;
import java.sql.Time;
import java.util.*;
/**
Разобраться почему при добавлении в базу данных id начинает увеличиваться не с последней id
*/
public class Main
{
    static final int secondsOfTheDay = 86400;
    static final int secondsOfAnHour = 3600;
    static final int secondsOfTheMinute = 60;
    private static boolean isSingular;
    private static final String path = "D:\\Для работы\\Projects\\Zepp Life\\SPORT.csv";
    private static final String fileInfo = "D:\\Для работы\\Projects\\Zepp Life\\info.xls";
    public static void addCitiesToDatabase(){
        SessionFactory factory = HibernateUtil.getSessionFactory();
        Session session = null;
        BufferedReader reader = null;
        try{
            try {
                reader = new BufferedReader(new FileReader("D:\\Для работы\\Projects\\Zepp Life\\Расстояния между городами России (Lite).txt"));
                String line = reader.readLine();
                while (reader.ready()) {
                    session = factory.getCurrentSession();
                    DistanceBetweenCities distanceBetweenCities = new DistanceBetweenCities();
                    line = reader.readLine();
                    String[] data = line.split(";");
                    System.out.println(data[0]+"|||"+data[1]+"|||"+data[2]);
                    distanceBetweenCities.setCityFrom(data[0]);
                    distanceBetweenCities.setCityTo(data[1]);
                    distanceBetweenCities.setDistance(Integer.parseInt(data[2]));
                    session.beginTransaction();
                    session.save(distanceBetweenCities);
                    session.getTransaction().commit();
                }
            } catch (IOException e) {
                System.out.println(e.getMessage());
            }
        }finally {
            factory.close();
            session.close();
        }
    }
    public static void showDataInConsole() throws FileNotFoundException{
        SessionFactory factory = HibernateUtil.getSessionFactory();
        Session session = null;
        BufferedReader reader = new BufferedReader(new FileReader(path));
        double commonPath=0.00;
        double pathOfTraining;
        int countOfTraining=1;
        Map<String,List<Integer>> infoTrainingsByDate = new LinkedHashMap<>();
        try{
            reader.readLine();//чтение линии для того, чтобы пропустить заголовки
            while(reader.ready()){
                /**
                 result[1]-дата + время начала тренировки
                 result[2]-время затраченное на тренировку в секундах
                 result[5]-дистанция тренировки
                 result[7]-калории за тренировку
                 */
                String[] result = reader.readLine().split(",");
                if(Integer.parseInt(result[0])!=9){
                    /**
                     date[0]-год
                     date[1]-месяц
                     date[2]-день
                     */
                    String[] date = result[1].substring(0,10).split("-");
                    /**
                     startExerciseTime[0]-часы начала тренировки по гринвичу, необходимо прибавить 4 чтобы время стало Самарским
                     startExerciseTime[1]-минуты
                     startExerciseTime[2]-секунды
                     */
                    String startExerciseTime = result[1].substring(11,19);
                    List<Integer> dataOfTraining = new ArrayList<>();
                    List<Integer> dataOfTraining2 = new ArrayList<>();
                    dataOfTraining.add(Integer.parseInt(result[2]));
                    dataOfTraining.add((int) Double.parseDouble(result[5]));
                    dataOfTraining.add((int) Double.parseDouble(result[7]));
                    int intermediateDistance = 0;
                    int intermediateTime = 0;
                    int intermediateCalories = 0;

                    try{
                        session = factory.getCurrentSession();
                        session.beginTransaction();
                        Training training = new Training();
                        training.setTrainingDate(Date.valueOf(result[1].substring(0,10)));
                        training.setTimeSeconds(Integer.parseInt(result[2]));
                        training.setDistance((int) Double.parseDouble(result[5]));
                        training.setCalorie((int) Double.parseDouble(result[7]));
                        training.setPace(Time.valueOf(calculateRunningPace((int) Double.parseDouble(result[5]), Integer.parseInt(result[2]))));
                        training.setStartExerciseTime(Time.valueOf(transformTimeToLocalTime(startExerciseTime)));
                        session.save(training);
                        session.getTransaction().commit();
                    }finally {
                        factory.close();
                        session.close();
                    }

                    try{
                        dataOfTraining2 = infoTrainingsByDate.get(date[2]+"."+date[1]+"."+date[0]);
                        intermediateTime = dataOfTraining2.get(0)+Integer.parseInt(result[2]);
                        intermediateDistance = dataOfTraining2.get(1)+(int) Double.parseDouble(result[5]);
                        intermediateCalories = dataOfTraining2.get(2)+(int) Double.parseDouble(result[7]);
                        dataOfTraining2.set(0,intermediateTime);
                        dataOfTraining2.set(1,intermediateDistance);
                        dataOfTraining2.set(2,intermediateCalories);
                        infoTrainingsByDate.put(date[2]+"."+date[1]+"."+date[0],dataOfTraining2);
                    }catch(NullPointerException e){
                        infoTrainingsByDate.put(date[2]+"."+date[1]+"."+date[0],dataOfTraining);
                    }
//                    pathOfTraining = Double.parseDouble(result[5])/1000;
//                    countOfTraining++;
//                    commonPath = (commonPath + pathOfTraining);
//                    System.out.println(String.format("%03d",countOfTraining)+"   "+date[2]+"/"+date[1]+"/"+date[0]+"   "+ Test.calculateTimeOfExercise(Integer.parseInt(result[2]))+
//                            "   "+String.format("%.2f",pathOfTraining)
//                            +"   "+String.format("%06.2f",commonPath));
                }
            }
        }catch(IOException exp){

        }
        for(Map.Entry<String,List<Integer>> entry: infoTrainingsByDate.entrySet()){
            System.out.println(String.format("%03d",(infoTrainingsByDate.size()-countOfTraining))+"  "+entry.getKey()+"    " + entry.getValue());
            countOfTraining++;
        }

    }
    public static void updateDatabase()throws FileNotFoundException{
        SessionFactory factory = HibernateUtil.getSessionFactory();
        Session session = null;
        BufferedReader reader = new BufferedReader(new FileReader(path));
        SortedMap<String,List<String>> infoTrainingsByDate = new TreeMap<>();
        SortedMap<String,List<String>> infoTrainingsByDateReverse = new TreeMap<>();
        int countNewTrainings=0;
        try{
            reader.readLine();//чтение линии для того, чтобы пропустить заголовки
            while(reader.ready()){
                /*
                 result[1]-дата + время начала тренировки
                 result[2]-время затраченное на тренировку в секундах
                 result[5]-дистанция тренировки
                 result[7]-калории за тренировку
                 */
                String[] result = reader.readLine().split(",");
                if(Integer.parseInt(result[0])!=9){
                    /**
                     date[0]-год
                     date[1]-месяц
                     date[2]-день
                     */
                    String[] date = result[1].substring(0,10).split("-");
                    /**
                     startExerciseTime[0]-часы начала тренировки по гринвичу, необходимо прибавить 4 чтобы время стало Самарским
                     startExerciseTime[1]-минуты
                     startExerciseTime[2]-секунды
                     */
                    String startExerciseTime = result[1].substring(11,19);
                    List<String> dataOfTraining = new ArrayList<>();
                    List<String> dataOfTraining2 = new ArrayList<>();
                    dataOfTraining.add(result[2]);
                    dataOfTraining.add(result[5]);
                    dataOfTraining.add(result[7]);
                    dataOfTraining.add(startExerciseTime);
                    try{
                        dataOfTraining2 = infoTrainingsByDate.get(result[1].substring(0,10));
                        dataOfTraining2.set(0,(String.valueOf(Integer.parseInt(dataOfTraining2.get(0))+Integer.parseInt(result[2]))));
                        dataOfTraining2.set(1,String.valueOf((int) Double.parseDouble(dataOfTraining2.get(1))+(int) Double.parseDouble(result[5])));
                        dataOfTraining2.set(2,String.valueOf((int) Double.parseDouble(dataOfTraining2.get(2))+(int) Double.parseDouble(result[7])));
                        dataOfTraining2.set(3,startExerciseTime);
                        infoTrainingsByDate.put(result[1].substring(0,10),dataOfTraining2);
                    }catch(NullPointerException e){
                        infoTrainingsByDate.put(result[1].substring(0,10),dataOfTraining);
                    }
                }
            }

        }catch(IOException exp){

        }

        try{
            List<String> inverseMap = new ArrayList<>(infoTrainingsByDate.keySet());
            Collections.reverse(inverseMap);
            for(String s:inverseMap){
                for(Map.Entry<String,List<String>> entry: infoTrainingsByDate.entrySet()){
                    if(entry.getKey().equals(s)){
                        infoTrainingsByDateReverse.put(entry.getKey(),entry.getValue());
                    }
                }
            }

            for(Map.Entry<String,List<String>> entry: infoTrainingsByDateReverse.entrySet()){
                List<String> dataOfTraining = new ArrayList<>();
                dataOfTraining = entry.getValue();
                String dateToCheck = "\'"+entry.getKey()+"\'";
                String hql = "select id from Training where trainingDate="+dateToCheck;
                session = factory.getCurrentSession();
                session.beginTransaction();
                org.hibernate.query.Query query = session.createQuery(hql);
                if(query.list().isEmpty()){
                    Training training = new Training();
                    training.setTrainingDate(Date.valueOf(entry.getKey()));
                    training.setTimeSeconds(Integer.parseInt(dataOfTraining.get(0)));
                    training.setDistance((int) Double.parseDouble(dataOfTraining.get(1)));
                    training.setCalorie((int) Double.parseDouble(dataOfTraining.get(2)));
                    training.setPace(Time.valueOf(calculateRunningPace((int) Double.parseDouble(dataOfTraining.get(1)), Integer.parseInt(dataOfTraining.get(0)))));
                    training.setStartExerciseTime(Time.valueOf(transformTimeToLocalTime(dataOfTraining.get(3))));
                    session.save(training);
                    session.getTransaction().commit();
                    System.out.println("Добавлена тренировка со следующими параметрами: \nдата: "+training.getTrainingDate()
                            +"; затрачено секунд: " + training.getTimeSeconds()+ "; дистанция: " + training.getDistance()
                            +"; затрачено калорий: " + training.getCalorie() + "; темп тренировки: "+training.getPace()
                            +"; время старта тренировки: " + training.getStartExerciseTime());
                    countNewTrainings++;
                }else{session.getTransaction().rollback();}
            }
            System.out.println("внесено в базу данных новых тренировок: "+countNewTrainings);//Количество внесенных тренировок
        }finally {
            factory.close();
            session.close();
        }
    }
    public static void createTableProducts() throws FileNotFoundException{
        SessionFactory factory = HibernateUtil.getSessionFactory();
        Session session = null;
        BufferedReader reader = new BufferedReader(new FileReader("D:\\Для работы\\Projects\\Zepp Life\\Таблица с калориями.txt"));
        try{
            try{
                reader.readLine();
                while(reader.ready()){
                    String[] data = reader.readLine().split(";");
                    session = factory.getCurrentSession();
                    session.beginTransaction();
                    Product product = new Product();
                    product.setName(data[0]);
                    product.setCalories(Integer.parseInt(data[1]));
                    product.setCaloriesPer(Double.parseDouble(data[2]));
                    product.setLitersOrKilograms(data[3]);
                    product.setSingularOrPlural(data[4]);
                    session.save(product);
                    session.getTransaction().commit();
                }
            }catch(IOException e){
                System.out.println(e.getMessage());
            }
        }finally {
            factory.close();
            session.close();
        }
    }
    public static HSSFWorkbook readWorkbook(String filename) {
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filename));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            return wb;
        }
        catch (Exception e) {
            return null;
        }
    }
    public static String calculateRunningPace(int distance, double time) {
        String pace=null;
        String min = "";
        String sec = "";
        double minute = (time/60)/distance*1000;
        double seconds = (minute-(int)minute)*60;
        if(minute>=0&&minute<10) {min = "0"+(int)minute;} else {min = String.valueOf((int)minute);}
        if(seconds>=0&&seconds<10) {sec = "0"+(int)seconds;} else {sec = String.valueOf((int)seconds);}
        pace="00:"+min+":"+sec;
        return pace;
    }
    public static String transformTimeToLocalTime(String time){
        String[] startExerciseTime= time.split(":");
        int hour = Integer.parseInt(startExerciseTime[0]);
        if(hour>=0&&hour<10){time = "0"+(hour+4)+":"+startExerciseTime[1]+":"+startExerciseTime[2];}
        else{time = (hour+4)+":"+startExerciseTime[1]+":"+startExerciseTime[2];}
        return time;
    }
    public static void writeWorkbook(HSSFWorkbook wb, String fileName) {
        try {
            FileOutputStream fileOut = new FileOutputStream(fileName);
            wb.write(fileOut);
            fileOut.close();
        }
        catch (Exception e) {
            //Обработка ошибки
        }
    }
    public static String calculateTimeOfExercise (int x){
        int interim=x;
        int days=0;
        int hours = 0;
        int minutes = 0;
        int seconds =0;
//        if(interim>=secondsOfTheDay){
//            days = interim/secondsOfTheDay;
//            interim = interim%secondsOfTheDay;
//        }
        if(interim>=secondsOfAnHour){
            hours = interim/secondsOfAnHour;
            interim = interim%secondsOfAnHour;
        }
        if(interim>=60){
            minutes = interim/secondsOfTheMinute;
            interim=interim%secondsOfTheMinute;
        }
        if(interim<60) seconds = interim;
//        return String.format("%02d",days)+" суток "+ String.format("%02d",hours) + ":" + String.format("%02d",minutes) + ":" + String.format("%02d",seconds);
        return String.format("%02d",hours) + ":" + String.format("%02d",minutes) + ":" + String.format("%02d",seconds);

    }
    public static String convertToRussianFormatDate (String date){
        String[] newFormat = date.split("-");
        return newFormat[2]+"."+newFormat[1]+"."+newFormat[0];
    }
    public static int randomWithRange(int min, int max)
    {
        int range = (max - min) + 1;
        return (int)(Math.random() * range) + min;
    }
    public static String declensionWord(String word, boolean isSingular,Case caseOfWord){
        Elements elements;
        String out = "";
        String url = "https://sklonili.ru/";
        try {
            Document doc = Jsoup.connect(url+word).userAgent("YaBrowser/22.7.0.1842").get();
            elements = doc.getElementsByAttributeValue("data-title","Склонение");
//            System.out.println(elements);
//            System.out.println(elements.size());
            if(elements.size()==0){
                out = word;
            }
            else{
                if(isSingular) {
//                    out = elements.get(1).text();
                    switch (caseOfWord){
                        case IMENIT:
                            out = elements.get(0).text();
                            break;
                        case RODIT:
                            out = elements.get(1).text();
                            break;
                        case DAT:
                            out = elements.get(2).text();
                            break;
                        case VINIT:
                            out = elements.get(3).text();
                            break;
                        case TVORIT:
                            out = elements.get(4).text();
                            break;
                        case PREDLOJ:
                            out = elements.get(5).text();
                            break;
                        default:
                            break;
                    }
                }
                else {out = elements.get(7).text();
                    switch (caseOfWord){
                        case IMENIT:
                            out = elements.get(6).text();
                            break;
                        case RODIT:
                            out = elements.get(7).text();
                            break;
                        case DAT:
                            out = elements.get(8).text();
                            break;
                        case VINIT:
                            out = elements.get(9).text();
                            break;
                        case TVORIT:
                            out = elements.get(10).text();
                            break;
                        case PREDLOJ:
                            out = elements.get(11).text();
                            break;
                        default:
                            break;
                    }
                }
            }
            elements.clear();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return out;
    }
    private static void refreshTrainingsData() {

        SessionFactory factory = HibernateUtil.getSessionFactory();
        Session session = null;
        try{
            HSSFWorkbook writeBook = new HSSFWorkbook();
            HSSFSheet writeSheet = writeBook.createSheet();
            String value = null;
            Training training;
            DistanceBetweenCities distanceBetweenCities;
            session = factory.getCurrentSession();
            session.beginTransaction();

            /**Расчет количества дней тренировок*/
            Query query = session.createQuery("select count(a.id) from Training a");
            value = String.valueOf(query.list().get(0));
            HSSFRow CommonNumberOFTrainingsRow = writeSheet.createRow(0);
            HSSFCell CommonNumberOFTrainingsCell = CommonNumberOFTrainingsRow.createCell(1);
            CommonNumberOFTrainingsCell.setCellValue(value);

            /**Внесение первой и последней даты тренировок*/
            query = session.createQuery("select a.trainingDate from Training a where id = 1");
            value = convertToRussianFormatDate(String.valueOf(query.list().get(0)))+"-";
            query = session.createQuery("select a.trainingDate from Training a where a.id in (select max(b.id) from Training b)");
            value = value+convertToRussianFormatDate(String.valueOf(query.list().get(0)));
            HSSFCell firstAndLastDatesOFTrainingsCell = CommonNumberOFTrainingsRow.createCell(5);
            firstAndLastDatesOFTrainingsCell.setCellValue(value);

            /**Расчет затраченного времени на тренировки*/
            query = session.createQuery("select sum(a.timeSeconds) from Training a");
            int x = (int)(long)query.list().get(0);
            value = calculateTimeOfExercise(x);
            HSSFRow CommonTimeRow = writeSheet.createRow(1);
            HSSFCell CommonTimeCell = CommonTimeRow.createCell(1);
            CommonTimeCell.setCellValue(value);


            /**Расчет общего количества километров*/
            query = session.createQuery("select sum(a.distance) from Training a");
            value = String.format("%.2f",((double)(int)(long)query.list().get(0))/1000) + " км";
            HSSFRow commonKilometersOFTrainingsRow = writeSheet.createRow(2);
            HSSFCell commonKilometersOFTrainingsCell = commonKilometersOFTrainingsRow.createCell(1);
            commonKilometersOFTrainingsCell.setCellValue(value);

            /**Информация с городами*/
            int dist1 = (int)(long)query.list().get(0)/1000;
            int dist2 = ((int)(long)query.list().get(0)/1000)-5;
            String hql = "select a from DistanceBetweenCities a where a.distance <= " + dist1+" AND a.distance >="+dist2;
            System.out.println(hql);
            query = session.createQuery(hql);
            System.out.println(query.list().get(0).getClass());
            distanceBetweenCities = (DistanceBetweenCities) query.list().get(randomWithRange(0,query.list().size()));
            value = "Именно на таком расстоянии друг от друга находятся \"" + distanceBetweenCities.getCityFrom()
                    + "\" и \"" + distanceBetweenCities.getCityTo()+"\"";
            HSSFRow citiesRow = writeSheet.createRow(3);
            HSSFCell citiesCell = citiesRow.createCell(1);
            citiesCell.setCellValue(value);


            /**Расчет общего количества затраченных калорий*/
            query = session.createQuery("select sum(a.calorie) from Training a");
            value = String.format("%,d",((int)(long)query.list().get(0))) + " килокалорий";
            HSSFRow commonCaloriesOfTrainingsRow = writeSheet.createRow(4);
            HSSFCell commonCaloriesOfTrainingsCell = commonCaloriesOfTrainingsRow.createCell(1);
            commonCaloriesOfTrainingsCell.setCellValue(value);
            value="";
            Product product;
            query = session.createQuery("select e.id from Product e");
            query = session.createQuery("select e from Product e where e.id="+query.list().get(Main.randomWithRange(0,query.list().size())));
            product = (Product)query.list().get(0);
            System.out.println(product.getName());
            if(product.getSingularOrPlural().equals("singular")) {isSingular = true;}
            else{isSingular = false;}
            if(product.getName().contains(" ")){
                String[] data = product.getName().split(" ");
                for (String datum : data) {
                    value = value + declensionWord(datum, isSingular, Case.RODIT) + " ";
                }
            }
            else{
                value = declensionWord(product.getName(),isSingular,Case.RODIT);
            }
            String newValue = "Столько же калорий содержится в "+(String.format("%.2f",39763/ product.getCaloriesPer())) +" "+
                    declensionWord(product.getLitersOrKilograms(), false,Case.PREDLOJ) +" " + value;
            System.out.println(newValue);
            HSSFCell infoCommonCaloriesOfTrainingsCell = commonCaloriesOfTrainingsRow.createCell(5);
            infoCommonCaloriesOfTrainingsCell.setCellValue(newValue);


            /**Расчет максимального количества затраченных калорий*/
            query = session.createQuery("select a from Training a where a.calorie in (select max(b.calorie) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow maxCaloriesOfTrainingsRow = writeSheet.createRow(7);
            HSSFCell maxCaloriesOfTrainingsCell = maxCaloriesOfTrainingsRow.createCell(1);
            maxCaloriesOfTrainingsCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMaxCalorieOfTrainingCell = maxCaloriesOfTrainingsRow.createCell(2);
            paceMaxCalorieOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMaxCalorieOfTrainingCell = maxCaloriesOfTrainingsRow.createCell(3);
            timeMaxCalorieOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMaxCalorieOfTrainingCell = maxCaloriesOfTrainingsRow.createCell(4);
            distanceMaxCalorieOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMaxCalorieOfTrainingCell = maxCaloriesOfTrainingsRow.createCell(5);
            dateMaxCalorieOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMaxCalorieOfTrainingCell = maxCaloriesOfTrainingsRow.createCell(6);
            startExcerciseTimeMaxCalorieOfTrainingCell.setCellValue(value);


            /**Расчет минимального количества затраченных калорий*/
            query = session.createQuery("select a from Training a where a.calorie in (select min(b.calorie) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow minCaloriesOfTrainingsRow = writeSheet.createRow(8);
            HSSFCell minCaloriesOfTrainingsCell = minCaloriesOfTrainingsRow.createCell(1);
            minCaloriesOfTrainingsCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMinCalorieOfTrainingCell = minCaloriesOfTrainingsRow.createCell(2);
            paceMinCalorieOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMinCalorieOfTrainingCell = minCaloriesOfTrainingsRow.createCell(3);
            timeMinCalorieOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMinCalorieOfTrainingCell = minCaloriesOfTrainingsRow.createCell(4);
            distanceMinCalorieOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMinCalorieOfTrainingCell = minCaloriesOfTrainingsRow.createCell(5);
            dateMinCalorieOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMinCalorieOfTrainingCell = minCaloriesOfTrainingsRow.createCell(6);
            startExcerciseTimeMinCalorieOfTrainingCell.setCellValue(value);

            /**Расчет тренировки с самым низким темпом*/
            query = session.createQuery("select a from Training a where a.pace in (select max(b.pace) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow maxPaceOfTrainingRow = writeSheet.createRow(9);
            HSSFCell maxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(1);
            maxPaceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMaxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(2);
            paceMaxPaceOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMaxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(3);
            timeMaxPaceOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMaxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(4);
            distanceMaxPaceOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMaxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(5);
            dateMaxPaceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMaxPaceOfTrainingCell = maxPaceOfTrainingRow.createCell(6);
            startExcerciseTimeMaxPaceOfTrainingCell.setCellValue(value);

            /**Расчет тренировки с самым лучшим темпом*/
            query = session.createQuery("select a from Training a where a.pace in (select min(b.pace) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow minPaceOfTrainingRow = writeSheet.createRow(10);
            HSSFCell minPaceOfTrainingCell = minPaceOfTrainingRow.createCell(1);
            minPaceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMinPaceOfTrainingCell = minPaceOfTrainingRow.createCell(2);
            paceMinPaceOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMinPaceOfTrainingCell = minPaceOfTrainingRow.createCell(3);
            timeMinPaceOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMinPaceOfTrainingCell = minPaceOfTrainingRow.createCell(4);
            distanceMinPaceOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMinPaceOfTrainingCell = minPaceOfTrainingRow.createCell(5);
            dateMinPaceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMinPaceOfTrainingCell = minPaceOfTrainingRow.createCell(6);
            startExcerciseTimeMinPaceOfTrainingCell.setCellValue(value);

            /**Расчет тренировки с самым коротким расстоянием*/
            query = session.createQuery("select a from Training a where a.distance in (select min(b.distance) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow minDistanceOfTrainingRow = writeSheet.createRow(11);
            HSSFCell minDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(1);
            minDistanceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMinDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(2);
            paceMinDistanceOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMinDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(3);
            timeMinDistanceOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMinDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(4);
            distanceMinDistanceOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMinDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(5);
            dateMinDistanceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMinDistanceOfTrainingCell = minDistanceOfTrainingRow.createCell(6);
            startExcerciseTimeMinDistanceOfTrainingCell.setCellValue(value);

            /**Расчет тренировки с самым длинным расстоянием*/
            query = session.createQuery("select a from Training a where a.distance in (select max(b.distance) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow maxDistanceOfTrainingRow = writeSheet.createRow(12);
            HSSFCell maxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(1);
            maxDistanceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceMaxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(2);
            paceMaxDistanceOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeMaxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(3);
            timeMaxDistanceOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceMaxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(4);
            distanceMaxDistanceOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateMaxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(5);
            dateMaxDistanceOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeMaxDistanceOfTrainingCell = maxDistanceOfTrainingRow.createCell(6);
            startExcerciseTimeMaxDistanceOfTrainingCell.setCellValue(value);

            /**Расчет самой ранней тренировки*/
            query = session.createQuery("select a from Training a where a.startExerciseTime in (select min(b.startExerciseTime) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow earlyStartExerciseTimeOfTrainingRow = writeSheet.createRow(13);
            HSSFCell earlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(1);
            earlyStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceEarlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(2);
            paceEarlyStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeEarlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(3);
            timeEarlyStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceEarlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(4);
            distanceEarlyStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateEarlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(5);
            dateEarlyStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeEarlyStartExerciseTimeOfTrainingCell = earlyStartExerciseTimeOfTrainingRow.createCell(6);
            startExcerciseTimeEarlyStartExerciseTimeOfTrainingCell.setCellValue(value);

            /**Расчет самой поздней тренировки*/
            query = session.createQuery("select a from Training a where a.startExerciseTime in (select max(b.startExerciseTime) from Training b)");
            training = (Training) query.list().get(0);
            value = String.valueOf(training.getCalorie());
            HSSFRow lateStartExerciseTimeOfTrainingRow = writeSheet.createRow(14);
            HSSFCell lateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(1);
            lateStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getPace());
            HSSFCell paceLateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(2);
            paceLateStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = calculateTimeOfExercise((int)(long)training.getTimeSeconds());
            HSSFCell timeLateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(3);
            timeLateStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.format("%.2f",((double)(int)(long)training.getDistance())/1000) + " км";
            HSSFCell distanceLateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(4);
            distanceLateStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = convertToRussianFormatDate(String.valueOf(training.getTrainingDate()));
            HSSFCell dateLateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(5);
            dateLateStartExerciseTimeOfTrainingCell.setCellValue(value);
            value = String.valueOf(training.getStartExerciseTime());
            HSSFCell startExcerciseTimeLateStartExerciseTimeOfTrainingCell = lateStartExerciseTimeOfTrainingRow.createCell(6);
            startExcerciseTimeLateStartExerciseTimeOfTrainingCell.setCellValue(value);


            Main.writeWorkbook(writeBook,fileInfo);
            session.getTransaction().commit();
        }finally {
            factory.close();
            session.close();
        }

    }
    public static void main(String[] args) {
//        updateDatabase();
        refreshTrainingsData();
    }


}
