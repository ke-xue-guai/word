package com.export.word;

/**
 * 单实现类
 *
 * @param <T>
 */
public abstract class Single<T> {

    private T instance;

    /**
     * 创建实例
     */
    public abstract T create();

    /**
     * 单例的方式获取实例
     * <p>
     * 达到优化的目的
     */
    public T get() {
        if (this.instance == null) {
            this.instance = create();
        }
        return this.instance;
    }

}
