import java.io.IOException;

import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {

        Scanner teclado = new Scanner(System.in);
        int c = teclado.nextInt();
        int n = teclado.nextInt();
        System.out.println(c%n);

    }

}