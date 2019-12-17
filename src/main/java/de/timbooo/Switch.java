package de.timbooo;

public class Switch {
    private final String name;
    private final int switchRow;

    public Switch(String name, int switchRow) {
        this.name = name;
        this.switchRow = switchRow;
    }

    public String getName() {
        return name;
    }

    public int getSwitchRow() {
        return switchRow;
    }

    @Override
    public String toString() {
        return name;
    }
}
