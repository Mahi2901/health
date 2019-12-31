package ArrayProgram;

import java.util.Arrays;
import java.util.Scanner;

class Arr {

    public void secLarge(int arr[]) {

        int large = arr[0], sec = arr[0];
        for (int i = 0; i < arr.length; i++) {
            if (arr[i] > large) {
                sec = large;
                large = arr[i];

            } else if (arr[i] > sec) {
                sec = arr[i];

            }
        }

        System.out.println("Second Large: " + sec);
        System.out.println("Large : " + large);

    }

    public void seperate(int arr[]) {
        int c = 0;
        for (int i = 0; i < arr.length; i++) {
            if (arr[i] != 0) {
                arr[c] = arr[i];
                c++;
            }
        }
        while (c < arr.length) {
            arr[c] = 0;
            c++;
        }
        System.out.println(Arrays.toString(arr));
    }

    public void sorting(int arr[]) {
        int temp;
        for (int i = 0; i < arr.length; i++) {
            for (int j = i + 1; j < arr.length; j++) {
                if (arr[i] > arr[j]) {
                    temp = arr[i];
                    arr[i] = arr[j];
                    arr[j] = temp;
                }
            }
        }
        for (int i = 0; i <= arr.length; i++) {
            System.out.println(arr[i]);
        }
    }

    public void duplicate(int arr[]) {
        try {
            System.out.println("\n");
            this.sorting(arr);
            System.out.println("\n");
            int[] temp = new int[arr.length];
            int j = 0;
            for (int i = 0; i < arr.length - 1; i++) {
                if (arr[i] != arr[i + 1]) {
                    temp[j++] = arr[i];
                }
            }

            temp[j++] = arr[arr.length - 1];
            // Changing original array  
            for (int i = 0; i < j; i++) {
                arr[i] = temp[i];
                System.out.print(arr[i] + " ");
            }

        } catch (ArrayIndexOutOfBoundsException e) {
            System.out.println(e);
        }
    }
}

public class SecondLarge {

    public static void main(String[] args) {
        Arr a = new Arr();
        Scanner sc = new Scanner(System.in);
        int arr[] = new int[5];

        System.out.println("Enter 5 Numbers = >");
        for (int i = 0; i <= 4; i++) {
            arr[i] = sc.nextInt();
        }

        /* a.secLarge(arr);
         System.out.println("*********************************************");
         a.seperate(arr);
         System.out.println("*********************************************");
         a.sorting(arr);
         System.out.println("*********************************************");
         */
        a.secLarge(arr);
    }
}
