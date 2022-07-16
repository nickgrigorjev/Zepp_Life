import org.hibernate.SessionFactory;
import org.hibernate.cfg.Configuration;

public class HibernateUtil {
    public static SessionFactory factory;
    private HibernateUtil(){}
public static synchronized SessionFactory getSessionFactory(){
        if (factory==null){
            factory =new Configuration().configure("hibernate.cfg.xml").
                    addAnnotatedClass(Training.class).
                    addAnnotatedClass(DistanceBetweenCities.class)
                    .addAnnotatedClass(Product.class).
                    buildSessionFactory();
        }
        return factory;
}
}
