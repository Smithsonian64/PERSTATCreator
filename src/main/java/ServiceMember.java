

import org.apache.poi.ss.usermodel.Cell;

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
    Cell endDateCellReference;

    ServiceMember(String name, String rank, String MOS, String orders, LocalDate startDate, LocalDate endDate, String taskForce, String status, Cell dateCellReference) {

        this.name = name;
        this.rank = rank;
        this.MOS = MOS;
        this.orders = orders;
        this.startDate = startDate;
        this.endDate = endDate;
        this.taskForce = taskForce;
        this.status = status;
        this.endDateCellReference = dateCellReference;

    }

}
