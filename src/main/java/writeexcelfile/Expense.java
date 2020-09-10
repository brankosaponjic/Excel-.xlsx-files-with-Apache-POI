package writeexcelfile;

import java.util.Date;

public class Expense {

    private int id;
    private String type;
    private double amount;
    private String currency;
    private Date date;

    public Expense(int id, String type, double amount, String currency, Date date) {
        this.id = id;
        this.type = type;
        this.amount = amount;
        this.currency = currency;
        this.date = date;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public double getAmount() {
        return amount;
    }

    public void setAmount(double amount) {
        this.amount = amount;
    }

    public String getCurrency() {
        return currency;
    }

    public void setCurrency(String currency) {
        this.currency = currency;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
