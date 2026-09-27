import java.util.Scanner;

public class Main {
    public static void main(String[] args) {

        Scanner s = new Scanner(System.in);

        while (s.hasNextDouble()) {
            double a = s.nextDouble();
            double r= Math.sqrt(3)/4.00;
            double area=r*a*a;
            System.out.printf("%.2f\n",area*1.6);

        }
    }
}
