import lombok.Data;

import javax.persistence.*;

@Data
@Entity
@Table(name = "distance_between_cities")
public class DistanceBetweenCities {
    @Id
    @SequenceGenerator(name = "distance_between_cities_seq")
    @GeneratedValue(strategy = GenerationType.SEQUENCE,generator = "distance_between_cities_seq")
    @Column(name = "id")
    private int id;
    @Column(name = "distance")
    private int distance;
    @Column(name = "city_from")
    private String cityFrom;
    @Column(name = "city_to")
    private String cityTo;

}
