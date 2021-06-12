

import java.time.LocalDate;

public class ServiceMember {

    String name;
    String rank;
    String MOS;
    String orders;
    LocalDate startDate;
    LocalDate endDate;
    String taskForce;
    String status;

    ServiceMember(String name, String rank, String MOS, String orders, LocalDate startDate, LocalDate endDate, String taskForce, String status) {

        this.name = name;
        this.rank = rank;
        this.MOS = MOS;
        this.orders = orders;
        this.startDate = startDate;
        this.endDate = endDate;
        this.taskForce = taskForce;
        this.status = status;

    }

}
