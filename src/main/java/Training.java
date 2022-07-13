import lombok.Data;
import javax.persistence.*;
import java.util.Calendar;
import java.util.Date;

@Data
@Entity
@Table(name = "trainings")
public class Training {
    @Id
    @SequenceGenerator(name = "tr_seq")
    @GeneratedValue(strategy = GenerationType.SEQUENCE,generator = "tr_seq")
    @Column(name = "id")
    private int id;
    @Column(name = "distance")
    private int distance;
    @Column(name = "calorie")
    private int calorie;
    @Column(name = "seconds")
    private int timeSeconds;
    @Column(name = "pace")
    private java.sql.Time pace;
    @Column(name = "training_date")
    private java.sql.Date trainingDate;
    @Column(name = "start_exercise_time")
    private java.sql.Time startExerciseTime;
}
