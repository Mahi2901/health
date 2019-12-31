/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package Mahendra;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

/**
 *
 * @author Student
 */
public class Factorial {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)throws IOException{ 
        // TODO code application logic here
         BufferedReader br =new BufferedReader( new InputStreamReader(System.in));
    int no,fact=1,i;
    System.out.print("Enter value of number==>");
    no=Integer.parseInt(br.readLine());
    
    for(i=1;i<=no;i++)
    {
         fact=fact*i;
    }
    System.out.println("Factorial = " + fact);
    }
    
    
}
