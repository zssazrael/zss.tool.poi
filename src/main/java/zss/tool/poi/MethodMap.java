package zss.tool.poi;

import java.lang.reflect.Method;
import java.util.TreeMap;

import zss.tool.Version;

@Version("2017.05.10")
class MethodMap extends TreeMap<String, Method> {
    private static final long serialVersionUID = 20170510221044L;

    private Class<?> type;

    public void setType(final Class<?> type) {
        if (type == null) {
            this.type = null;
            clear();
            return;
        }
        if (type == this.type) {
            return;
        }
        this.type = type;
        for (Method method : type.getMethods()) {
            if (method.getReturnType() == Void.TYPE) {
                continue;
            }
            if (method.getParameterTypes().length > 0) {
                continue;
            }
            put(method.getName(), method);
        }
    }

    public Method getMethod(String name) {
        if (name == null) {
            return null;
        }
        Method method = get("get".concat(name));
        if (method == null) {
            method = get("is".concat(name));
        }
        if (method == null) {
            method = get(name);
        }
        return method;
    }
}
