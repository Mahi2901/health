/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package Mahendra;

/**
 *
 * @author Student
 */
public class pra5 {
     /**
     * @param args the command line arguments
     */
    public static void main(String[] args){ 
        // TODO code application logic here
   long sum;
   long a=0,b=1;
   long n=100;
   int i;
   
   for(i=1;i<=n;i++)
   {
       sum=a+i;
       System.out.println(sum + "");
      a=b;
      b=sum;
      
   } 
   
    
    
}
}
