package org.example.model;

public class MemberVO {
    private String name;
    private int age;
    private String birthdate;
    private String phone;
    private String address;
    private boolean isMarried;

    public MemberVO(String name, int age, String birthdate, String phone, String address, boolean isMarried) {
        this.name = name;
        this.age = age;
        this.birthdate = birthdate;
        this.phone = phone;
        this.address = address;
        this.isMarried = isMarried;
    }

    public String getName() {
        return name;
    }

    public int getAge() {
        return age;
    }

    public String getBirthdate() {
        return birthdate;
    }

    public String getPhone() {
        return phone;
    }

    public String getAddress() {
        return address;
    }

    public boolean isMarried() {
        return isMarried;
    }

    @Override
    public String toString() {
        return "MemberVO{ " +
                "name='" + name + '\'' +
                ", age=" + age +
                ", birthdate='" + birthdate + '\'' +
                ", phone='" + phone + '\'' +
                ", address='" + address + '\'' +
                ", isMarried=" + isMarried +
                " }";
    }
}
