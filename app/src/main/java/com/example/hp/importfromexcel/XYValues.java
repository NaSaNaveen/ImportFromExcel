package com.example.hp.importfromexcel;

/**
 * Created by HP on 02-Oct-17.
 */

public class XYValues
{
    private double x;
    private double y;

    public double getX() {
        return x;
    }

    public void setX(double x) {
        this.x = x;
    }

    public double getY() {
        return y;
    }

    public void setY(double y) {
        this.y = y;
    }

    public XYValues(double x, double y) {

        this.x = x;
        this.y = y;
    }
}
