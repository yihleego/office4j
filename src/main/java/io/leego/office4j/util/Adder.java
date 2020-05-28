package io.leego.office4j.util;

/**
 * @author Yihleego
 */
public class Adder {
    private int v;

    public Adder() {
        this.v = 0;
    }

    public Adder(int v) {
        this.v = v;
    }

    public int addAndGet(int delta) {
        v += delta;
        return v;
    }

    public int getAndGet(int delta) {
        int o = v;
        v += delta;
        return o;
    }

    public int incrementAndGet() {
        return ++v;
    }

    public int getAndIncrement() {
        return v++;
    }

    public int get() {
        return v;
    }

}
