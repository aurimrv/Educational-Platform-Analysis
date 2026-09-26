import java.util.Scanner;
public class Main {
    public static void main(String[] args) {
        Scanner s = new Scanner(System.in);
        String input = s.nextLine();

        System.out.println(input.toLowerCase().contains("zelda") ? "Link Bolado" : "Link Tranquilo");
    }
}
