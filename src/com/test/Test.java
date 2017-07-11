package com.test;

public class Test {
	
	public static void main(String[] args) {
		String str = "sreenivasulu.s@gavstech.com";
		String split_first = str.substring(0,str.indexOf("@"));
		String split_second = str.substring(str.indexOf("@")+1);
		String[] parts = split_second.split("\\.");
		String part1 = parts[0]; 
		String part2 = parts[1];
		
		System.out.println(part1);
	}
	
}
