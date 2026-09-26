import java.util.Scanner;

public class Main {

	public static void main(String[] args) {
		
		Scanner sc = new Scanner(System.in);
		
		long x = sc.nextLong();
		long y = sc.nextLong();		
		
		long sum = (x + y) * (y - x + 1) / 2;		
		System.out.println(sum);		
		
		sc.close();
	}
}
