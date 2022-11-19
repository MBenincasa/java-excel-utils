package samples.convertObjectsToExcelFile;

public class Employee {

    private String name;
    private String lastName;
    private Integer age;
    private Double salary;

    public Employee(String name, String lastName, Integer age, Double salary) {
        this.name = name;
        this.lastName = lastName;
        this.age = age;
        this.salary = salary;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "name='" + name + '\'' +
                ", lastName='" + lastName + '\'' +
                ", age=" + age +
                ", salary=" + salary +
                '}';
    }
}
