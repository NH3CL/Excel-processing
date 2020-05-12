package com.example.test;

public class Entry {

    String id;
    String month;
    String phone;

    String mail;

    String AP;
    String BN;
    String BI;
    public Entry(String id, String month, String phone, String mail, String AP, String BN, String BI) {
        this.id = id;
        this.month = month;
        this.phone = phone;
        this.mail = mail;
        this.AP = AP;
        this.BN = BN;
        this.BI = BI;
    }

    public Entry() {

    }

    @Override
    public String toString() {
        return "Entry{" +
                "id=" + id +
                ", month='" + month + '\'' +
                ", phone='" + phone + '\'' +
                ", mail='" + mail + '\'' +
                ", AP='" + AP + '\'' +
                ", BN='" + BN + '\'' +
                ", BI='" + BI + '\'' +
                '}';
    }

    public void setId(String id) {
        this.id = id;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public void setMail(String mail) {
        this.mail = mail;
    }

    public void setAP(String AP) {
        this.AP = AP;
    }

    public void setBN(String BN) {
        this.BN = BN;
    }

    public void setBI(String BI) {
        this.BI = BI;
    }

    public String getMonth() {
        return month;
    }

    public String getPhone() {
        return phone;
    }

    public String getId() {
        return id;
    }

    public String getMail() {
        return mail;
    }

    public String getAP() {
        return AP;
    }

    public String getBN() {
        return BN;
    }

    public String getBI() {
        return BI;
    }
}
