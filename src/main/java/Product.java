import lombok.Data;

import javax.persistence.*;

@Data
@Entity
@Table(name = "products_with_calories")
public class Product {
    @Id
    @SequenceGenerator(name = "product_seq")
    @GeneratedValue(strategy = GenerationType.SEQUENCE,generator = "product_seq")
    @Column(name = "id")
    private int id;
    @Column(name = "product_name")
    private String name;
    @Column(name = "singular_or_plural")
    private String singularOrPlural;
    @Column(name = "calories_per")
    private double caloriesPer;
    @Column(name = "liters_or_kilograms")
    private String litersOrKilograms;
    @Column(name = "calories")
    private int calories;

}
