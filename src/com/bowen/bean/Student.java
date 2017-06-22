package com.bowen.bean;

public class Student {
  private long id;
  private String name;
  private int age;
  private String sex;
  private String address;
  
public Student(long id, String name, int age, String sex, String address) {
	super();
	this.id = id;
	this.name = name;
	this.age = age;
	this.sex = sex;
	this.address = address;
}

public long getId() {
	return id;
}

public void setId(long id) {
	this.id = id;
}

public String getName() {
	return name;
}

public void setName(String name) {
	this.name = name;
}

public int getAge() {
	return age;
}

public void setAge(int age) {
	this.age = age;
}

public String getSex() {
	return sex;
}

public void setSex(String sex) {
	this.sex = sex;
}

public String getAddress() {
	return address;
}

public void setAddress(String address) {
	this.address = address;
}

@Override
public String toString() {
	return "Student [id=" + id + ", name=" + name + ", age=" + age + ", sex=" + sex + ", address=" + address + "]";
}
  
  
  
}
